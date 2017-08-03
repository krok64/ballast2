"""
версия 1.00
Расчет балластировки и подсадки
версия 1.01 16.03.2016
Добавить вывод профилей и планов в автокад масштаб для линии плана 1:1000 для профиля 1:2000 горизонтальный, 1:200 вертикальный
версия 1.02 22.03.2016
Полностью переделан шаблон, улучшена точность нахождения дельта ф, новая трасса задается двумя массивами (x2d y2d) и потом пересчитывается в xd
"""

import re
from math import pi, cos, ceil, exp
import os
import sys
import configparser
import array

import numpy as np
import matplotlib.pyplot as plt
import win32com.client
from PIL import Image
import pymorphy2
from pyautocad import Autocad, APoint

from mdvlib.tpcalc import get_k_n, GetTExp
from mdvlib.mso import word_table_fill
from mdvlib.util import str_to_float, str_to_arr_rus_float

def poly_num(x,y,num):
#вернуть сглаженный полиномом num степени массив y(x)
    y_s=np.poly1d(np.polyfit(x,y,num))
    return [y_s(x1) for x1 in x]
    
def Izg_Napr(x,y,D,E):
#изгибающее напряжение в трубопроводе по 5 точкам, по графику y(x), 
#E-модуль упругости, D-внешний диаметр. 2 первые и 2 последние точки не считаем
#возвращем в МПа
    Ksi=[]
    for i in range(2,len(x)-2):
        y_c=np.polyfit(x[i-2:i+3],y[i-2:i+3],2)
        r=1/(2*y_c[0])
        Ksi.append(E*D/(2*r)/10**6)
    return Ksi

def find_point(x, y, x_cur):
    #возвращает точку с прямой x y которая имеет координату x_cur
    for i in range(len(x)):
            if x_cur > x[i]:
                continue
            break
#    print(y[i-1], x[i-1], y[i], x[i], x_cur)
    return(Inter(y[i-1], x[i-1], y[i], x[i], x_cur))

def MaxDY(x0,L,x,y):
#расчет расстояния (по y) от прямой через точки y(x0) и y(x0+L) до минимальной точке на кривой y(x)
    x_lin=[x0,x0+L]
#    y_s=np.poly1d(np.polyfit(x,y,6))
    y_lin=[find_point(x,y,x0),find_point(x,y,x0+L)]
    y_s_lin=np.poly1d(np.polyfit(x_lin,y_lin,1))
    min_y=min(y)
    for i in range(0,len(x)):
        if (y[i]-min_y)<0.0001:
            min_x=x[i]
    return y_s_lin(min_x)-min_y

def CalcIntersection(x,y,y0):
# найти все пересечения графика y=f(x) c прямой y=y0
    TOP,BOT,EQ=range(3)
    ret=[]
    if y[0]>y0:
        pos=TOP
    elif y[0]<y0:
        pos=BOT
    else:
        pos=EQ
        ret.append(x[0])

    for i in range(1, len(y)):
        if y[i]>y0:
            if pos==BOT:
                x0=Inter(x[i-1],y[i-1],x[i],y[i],y0)
                ret.append(x0)
            pos=TOP
        elif y[i]<y0:
            if pos==TOP:
                x0=Inter(x[i-1],y[i-1],x[i],y[i],y0)
                ret.append(x0)
            pos=BOT
        else:
            pos=EQ
            ret.append(x[i])
    return ret

def Inter(x0,y0,x1,y1,y):
#найти такое x которое соответствует y   
    dx=x1-x0
    dy=y1-y0
    return x0+dx*(y-y0)/dy
    
def pline_to_acad(acad, x, y, x0, y0, xscale, yscale):
    # рисует полилинию по точкам x y в автокаде x0 y0 - начальная точка отрисовки xscale yscale - масштаб отрисовки (1:1000 = 1; 1:2000 = 0.5; 1:200 = 5)
    if len(x) != len (y):
        raise Exception("x and y must be same dimensions")
    points = array.array('d')
    for i in range(len(x)):
        points.append(x0+x[i]*xscale) 
        points.append(y0+y[i]*yscale) 
    #print(points)
    pl = acad.model.AddLightWeightPolyline(points)
    return pl
    
#настройка русского шрифта в matplotlib
from matplotlib import rc
rc('font',**{'family':'verdana'})
rc('text.latex',unicode=True)
rc('text.latex',preamble=r'\usepackage[utf8]{inputenc}')
rc('text.latex',preamble=r'\usepackage[russian]{babel}')

config = configparser.ConfigParser()


for filename in sys.argv[1:]:

    config.read(filename, encoding='utf-8')

    E = 2.06*10**11  #модуль упругости стали Па
    g = 9.81         #ускорение свободного падения м/с2
    Ro_st = 7850     #плотность стали кг/м3
    Ro_vod = 1000    #плотность воды кг/м3
    Ro_bet = 2400    #плотность бетона кг/м3 из ТУ на УтО
    n_b = 0.9        #к-т надежности по нагрузке 0,9 бетонные грузы, 1-чугунные грузы
    alpha = 0.12*10**(-4) #коэффициент линейного расширения стали 1/град
    
    D_n = str_to_float(config["Pipe"]["D"])/1000
    s_st = str_to_float(config["Pipe"]["s"])/1000
    sigma_t = str_to_float(config["Pipe"]["sigma_t"]) * 10**6
    p = str_to_float(config["Pipe"]["p_r"]) * 10**6
    p_n = str_to_float(config["Pipe"]["p_n"]) * 10**6
    p_k = str_to_float(config["Pipe"]["p_k"]) * 10**6
    x0 = str_to_float(config["Land"]["x0"])
    x1 = str_to_float(config["Land"]["x1"])
    x_kc = str_to_float(config["Pipe"]["X_kc"])
    l_kc = str_to_float(config["Pipe"]["L_kc"])
    L = x1 - x0
    m = str_to_float(config["Pipe"]["m"])
    k_n_v = str_to_float(config["Ballast"]["k_n_v"])
    x0_ballast = str_to_float(config["Ballast"]["x0"])
    x1_ballast = str_to_float(config["Ballast"]["x1"])
    L_ballast = x1_ballast - x0_ballast
    V_gruza = str_to_float(config["Ballast"]["V_gruza"])

    p_x = (p_n**2-(p_n**2-p_k**2)*x_kc/l_kc)**0.5
    #температуры газа в начале и в конце зимой и летом
    t_zima_n = 25 
    t_zima_k = 0
    t_leto_n = 30
    t_leto_k = 10
    
    #температуры трубы при ремонте зимой и летом
    t_rem_leto = 20
    t_rem_zima = -15
    
    t_exp_zima = GetTExp(t_zima_n, t_zima_k, l_kc, x_kc)
    t_exp_leto = GetTExp(t_leto_n, t_leto_k, l_kc, x_kc)
    
    delta_t_max = t_exp_leto - t_rem_zima
    delta_t_min = t_exp_zima - t_rem_leto

    #координата x
    xd=str_to_arr_rus_float(config["Coord"]["x"])
    #отметка трубы
    yd=str_to_arr_rus_float(config["Coord"]["y"])
    #отметка y обваловки
    y_obv=str_to_arr_rus_float(config["Coord"]["y_obv"])
    #отметка уровня земли
    ld=str_to_arr_rus_float(config["Coord"]["y_land"])
    #координата земли в горизонтальной плоскости
    zd=str_to_arr_rus_float(config["Coord"]["z"])
    #отметка переуложенной трубы
    y2d=str_to_arr_rus_float(config["Coord"]["y2"])
    #отметка земли для переуложенной трубы
    x2d=str_to_arr_rus_float(config["Coord"]["x2"])

    if (len(xd) != len(yd) or len(yd) != len(ld) or len(ld) != len(zd) or len(zd) != len(y_obv)):
        print("разная длина исходных данных: xd=%d yd=%d ld=%d zd=%d y_obv=%d" % (len(xd), len(yd), len(ld), len(zd), len(y_obv)))
        raise Exception("Invalid data")
    if (len(y2d) != len(x2d)):        
        print("разная длина исходных данных: xd=%d y2d=%d" % (len(x2d), len(y2d)))
        raise Exception("Invalid data")

    #пересчитываем y2d с x2d на xd
    a=[]
    a.append(y2d[0])
    for i in range(1, len(xd)):
#    for i in range(1, 10):
        my_x = xd[i]
#       print("my_x=%.2f" % my_x)
        for j in range(len(x2d)):
#            print("x2d=%.2f" % x2d[j])
            if my_x > x2d[j]:
                continue
            break
#        print(y2d[j-1], x2d[j-1], y2d[j], x2d[j], my_x)
        a.append(Inter(y2d[j-1], x2d[j-1], y2d[j], x2d[j], my_x))

    y2d = a
#    for i in range(len(y2d)):
#    for i in range(10):
#        print("x=%d y=%.2f" % (xd[i], y2d[i]), end=' ')
#    break

    #отметка воды
    vd=str_to_float(config["Land"]["y_vod"])

    r_sr=0.5*(D_n-s_st) #средний радиус сечения ТП
    J=pi*r_sr**3*s_st #осевой момент инерции сечения трубы
    W=pi*r_sr**2*s_st #осевой момент сопротивления сечения трубы
    F=2*pi*r_sr*s_st  #площадь сечения стенки трубы, м2
#находим координаты х пересечения верха переуложенной трубы с уровнем воды
    ballast_points=CalcIntersection(xd,y2d,vd)
    s_ballast=""
    for i in ballast_points:
        s_ballast+=" %.2f" % (i)
    
    #рисунок 1
    y_glad=poly_num(xd,yd,11)
    plt.figure(figsize=(13, 11))
    plt.plot(xd,y_obv,label="1", color="k")
    plt.plot(xd,ld,label="2", color="#964B00")
#вода - показываем ее по балластировке
    xb=[]
    xb.append(x0_ballast)
    xb.append(x1_ballast)
    yb=[]
    yb.append(vd)
    yb.append(vd)
    if config["Land"].getboolean("show_water", fallback=True) == False:
        plt.plot(xb,yb,label="3", linestyle = '-.', color="w")
    else:
        plt.plot(xb,yb,label="3", linestyle = '-.', color="b")
    plt.plot(xd,yd,label="4", linewidth = 2, color = "r")
    plt.plot(xd,y_glad,label="5", linestyle = '--', linewidth = 2, color="g")
#    plt.plot(xd,y2d,label="6", linewidth = 2, color = "b")
#линия балластировки    
    # xb=[]
    # xb.append(x0_ballast)
    # xb.append(x0_ballast+L_ballast)
    # yb=[]
    # yb.append(vd-1)
    # yb.append(vd-1)
    # plt.plot(xb,yb,label="7", linestyle = '-.', color="r")
    plt.xlabel('x, м')
    plt.ylabel('y, м')
    plt.legend(loc='lower right')
    
    plt.grid()
    plt.savefig(filename+"fig1.png")
    plt.clf()
    
    #поворот и обрезка первого рисунка
    im = Image.open(filename+"fig1.png")
    im = im.rotate(90)
    im = im.crop((190,0,1140,1100))
    im.save(filename+"fig1.png")

    #расчет f0 и f1
    f0=MaxDY(x0,L,xd,yd)
    f1=MaxDY(x0,L,xd,y2d)
            
    #рисунок 2
    z_glad=poly_num(xd,zd,6)
    plt.figure(figsize=(9, 4.5))
    plt.plot(xd,zd,label="1", linewidth = 2)
    plt.plot(xd,z_glad,label="2", linestyle = '--', linewidth = 2)
    plt.xlabel('x, м')
    plt.ylabel('z, м')
    plt.legend(loc='lower right')
    plt.grid()
    plt.savefig(filename+"fig2.png")
    plt.clf()

    #рисунок 3
    Ksi_y=Izg_Napr(xd,yd,D_n,E)
    Ksi_y_glad=Izg_Napr(xd,y_glad,D_n,E)
    plt.plot(xd[2:-2],Ksi_y,label="1", linewidth = 2)
    plt.plot(xd[2:-2],Ksi_y_glad,label="2", linestyle = '--', linewidth = 2)
    plt.xlabel('x, м')
    plt.ylabel(r'$\sigma$, МПа')
    plt.legend()
    plt.grid()
    plt.savefig(filename+"fig3.png")
    plt.clf()

    #рисунок 4
    Ksi_z=Izg_Napr(xd,zd,D_n,E) 
    Ksi_z_glad=Izg_Napr(xd,z_glad,D_n,E)
    plt.plot(xd[2:-2],Ksi_z,label="1", linewidth = 2)
    plt.plot(xd[2:-2],Ksi_z_glad,label="2", linestyle = '--', linewidth = 2)
    plt.xlabel('x, м')
    plt.ylabel(r'$\sigma$, МПа')
    plt.legend()
    plt.grid()
    plt.savefig(filename+"fig4.png")
    plt.clf()

    #рисунок 5
    Ksi_y2=Izg_Napr(xd,y2d,D_n,E)
    Ksi_px=[]
    for i in xd[2:-2]:
        Ksi_px.append((-2*pi**2/L**2*(f1-f0)*E*J*cos(2*pi*(i-x0)/L)/W)/10**6)
    Ksi_sum=[]
    Ksi_px_vert=[]
    for i in range(0, len(Ksi_y2)):
        Ksi_sum.append(((Ksi_y2[i]+Ksi_px[i])**2+Ksi_z_glad[i]**2)**0.5)
        Ksi_px_vert.append(Ksi_y2[i]+Ksi_px[i])
    plt.plot(xd[2:-2],Ksi_px_vert,label="1", linewidth = 2)
    plt.plot(xd[2:-2],Ksi_z_glad,label="2", linestyle = '--', linewidth = 2)
    plt.plot(xd[2:-2],Ksi_sum,label="3", linestyle = '-.', linewidth = 2)
    plt.xlabel('x, м')
    plt.ylabel(r'$\sigma$, МПа')
    plt.legend()
    plt.grid()
    plt.savefig(filename+"fig5.png")
    #plt.show()
    plt.cla()
    #print (W)
    #print (Ksi_px)    
    #print (Ksi_sum)    

    delta_f=f1-f0
    q_tr=pi*(D_n**2-(D_n-2*s_st)**2)/4*Ro_st*g
    q_iz=2*pi**4*((delta_f)*E*J+0.0938*(f1**3-f0**3)*E*F)/L**4
    k_n=get_k_n(D_n,p)
    Ksi_N=pi**2*(f1**2-f0**2)*E/(4*L**2)
    D_vn=D_n-2*s_st
    Ksi_kc_a=p_x*D_vn/(2*s_st)
    Ksi_kc_b=p*D_vn/(2*s_st)
# максимальные напряжение на переукладываемом участке чтобы отсечь напряжения которые не входят в переукладку
# находим номер первой отметки Х которая попадает под переукладку x>x0
    for i in range(len(xd)):
        if xd[i]>x0:
            break
# находим номер последней отметки Х которая попадает под переукладку x<x0+L
    for j in range(len(xd)-1,0,-1):
        if xd[j]<x0+L:
            break
    #print(i,j)
            
    Ksi_s_max=max(abs(x) for x in Ksi_sum[i-1:j+1])
    Ksi_pr_n_a=Ksi_N+0.3*Ksi_kc_a-alpha*delta_t_max*E-Ksi_s_max*10**6
    Ksi_pr_n_b=Ksi_N+0.3*Ksi_kc_b-alpha*delta_t_min*E+Ksi_s_max*10**6
    psi3=(1-0.75*(Ksi_kc_a/(m*sigma_t/(0.9*k_n)))**2)**0.5-0.5*Ksi_kc_a/(m*sigma_t/(0.9*k_n))
    Usl_prochnosti=m*sigma_t/(0.9*k_n)

    n_gruz=ceil((pi/4*k_n_v*D_n**2*g*Ro_vod-q_tr+k_n_v*q_iz)*L_ballast/(n_b*(Ro_bet-k_n_v*Ro_vod)*V_gruza*g))
    l_shag=L_ballast/n_gruz
    
    L_op_shag=(12*W*sigma_t/(q_tr*k_n))**0.5
    sigma_i = max(abs(x) for x in Ksi_y_glad)

    plt.title('Результаты расчетов')
    plt.ylim( -130, 190 ) 
    plt.xlim( 0, 400 ) 
    plt.text(50,180,r'$D_n=%.2fм, \delta=%.4fм, r_{средний}=%0.3fм, P_{раб}=%.1fМПа, P_x=%.1fМПа, R_2^н=%dМПа$' % (D_n, s_st, r_sr, p/10**6, p_x/10**6, sigma_t/10**6))
    plt.text(50,160,r'Участок переукладки: от %d м до %d м (длина %d м)' % (x0,x0+L,L))
    plt.text(50,140,r'$f_0=%0.2f м, f_1=%0.2f м, \Delta f=%0.2f м$' % (f0,f1, delta_f))
    plt.text(50,120,r'$|\sigma_и|=%0.2f МПа, |\sigma_{s.max}|=%0.2fМПа$' % (sigma_i, Ksi_s_max ))
    plt.text(50,100,r'$EF=%e Н, EJ=%e Нм^2$' % (E*F, E*J))
    plt.text(50,80,r'$q_{тр}=%.2f Н/м, q_{из}=%.2f Н/м$' % (q_tr, q_iz))
    plt.text(50,60,r'$k_н=%0.3f, \sigma_N=%0.2f МПа$' % (k_n, Ksi_N/10**6))
    plt.text(50,40,r'$\sigma_{кц(а)}^н=%0.2f МПа, \sigma_{кц(б)}^н=%0.2f МПа$' % (Ksi_kc_a/10**6, Ksi_kc_b/10**6))
    plt.text(50,20,r'$\sigma_{пр(а)}^н=%0.2f МПа, \sigma_{пр(б)}^н=%0.2f МПа$' % (Ksi_pr_n_a/10**6, Ksi_pr_n_b/10**6))
    plt.text(50,00,r'$\psi_3=%.2f$' % (psi3))
    plt.text(50,-20,r'$\psi_3\frac{m}{0.9k_н}R_2^н =%.2f МПа, \frac{m}{0.9k_н}R_2^н =%.2f МПа$' % (psi3*Usl_prochnosti/10**6, Usl_prochnosti/10**6))
    plt.text(50,-40,r'Число грузов $n_{У}=%d$, шаг установки грузов=%.2fм' % (n_gruz, l_shag))
    plt.text(50,-60,r'Участок балластировки: L=%d м (x = %d - %d)' % (L_ballast, x0_ballast, x0_ballast+L_ballast))
    plt.text(50,-80,r'Шаг опорных перемычек: %.1f м' % (L_op_shag))
    plt.text(50,-100,r'Уровень воды: %.2f м' % (vd))
    plt.text(50,-120,r'Пересечения трубы с водой: %s' % (s_ballast))
    plt.axis('off')
    plt.savefig(filename+"fig6.png")
    
    if config["Project"].getboolean("no_word", fallback=False):
        continue

    wordapp = win32com.client.Dispatch("Word.Application") # Create new Word Object
    worddoc = wordapp.Documents.Open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Балластировка и переукладка ШАБЛОН.docx")) # Create new Document Object
    worddoc.SaveAs2(os.path.abspath(filename)+".docx")
    
    #штамп
    try:
        worddoc.CustomDocumentProperties("Shifr").Value = config["Project"].get("shifr", fallback="Шифр проекта")
        worddoc.CustomDocumentProperties("Descr").Value = config["Project"].get("descr", fallback="Описание проекта")
        worddoc.CustomDocumentProperties("L_vskr").Value = "%d" % str_to_float(config["Land"]["L_vskr"])
        worddoc.CustomDocumentProperties("L_pods").Value = "%d" % L
        worddoc.CustomDocumentProperties("UBO_text").Value = "%s-%s (%d шт. шаг %.2fм) Расстояние пригрузки - %dм" % (config["Ballast"]["gruz_name"], config["Pipe"]["D"], n_gruz, l_shag, L_ballast )
    except:
        pass
    #Обновляем поля
    for aStory in worddoc.StoryRanges:
        for aField in aStory.Fields:
            aField.Update()
    worddoc.Fields.Update
    # Если у Вас есть поля в шейпах (например в надписях) то нужно обновить и их
    for v in worddoc.Shapes:
        if v.TextFrame.HasText:
            v.TextFrame.TextRange.Fields.Update()
            
    #Заполняем табличку исходных данных (таблица 1)
    nums = [i+1 for i in range(len(xd))]
    data = (nums, xd, yd, y_obv, ld, zd)
    word_table_fill(data, ("%d", "%.2f", "%.2f", "%.2f", "%.2f", "%.2f"), wordapp, worddoc, "indata", "Продолжение таблицы 1", True)
    
#     #Заполняем табличку расчетный данных (таблица 2)
#     y2d_dn = [y2d[i] - D_n for i in range(len(xd))]
#     yd_y2d = [yd[i] - y2d[i] for i in range(len(xd))]
#     data = (nums, xd, y2d, y2d_dn, yd_y2d)
# #    word_table_fill(data, ("%d", "%.2f", "%.2f", "%.2f", "%.2f"), wordapp, worddoc, "outdata", "Продолжение таблицы 2", True)

    morph = pymorphy2.MorphAnalyzer()
    kompl = morph.parse('комплект')[0]
    #Список полей из вордовского файла и значений которые они должны принимать
    doclist=[
        { "names": ("gruz","gruz2","gruz3"), "value": config["Ballast"]["gruz_name"] },
        { "names": ("lpumg", "lpumg2", "lpumg3", "lpumg4"), "value": config["Project"]["lpu"] },
        { "names": ("km", "km2", "km3"), "value": config["Project"]["km"] },
        { "names": ("tp_name","tp_name2", "tp_name3"), "value": config["Project"]["name"]},
        { "names": ("D", "D2", "D3", "D4"), "value": config["Pipe"]["D"]},
        { "names": ("s",), "value": config["Pipe"]["s"]},
        { "names": ("sigma_t","sigma_t2"), "value": config["Pipe"]["sigma_t"]},
        { "names": ("sigma_v",), "value": config["Pipe"]["sigma_v"]},
        { "names": ("y_voda",), "value": config["Land"]["y_vod"]},
        { "names": ("ubo","ubo2","ubo3"), "value": n_gruz },
        { "names": ("ubo_komp","ubo_komp2","ubo_komp3"), "value": kompl.make_agree_with_number(n_gruz).word },
        { "names": ("shag_ubo","shag_ubo2","shag_ubo3"), "value": "%.2f" % l_shag},
        { "names": ("x0_ball","x0_ball2","x0_ball3"), "value": "%d" % x0_ballast},
        { "names": ("x1_ball","x1_ball2","x1_ball3"), "value": "%d" % x1_ballast},
        { "names": ("sigma_i",), "value": "%0.2f" % sigma_i},
        { "names": ("L","L2","L3"), "value": "%d" % L},
        { "names": ("delta_f","delta_f2",), "value": "%.1f" % delta_f},
        { "names": ("L_vskr","L_vskr2"), "value": "%d" % str_to_float(config["Land"]["L_vskr"])},
        { "names": ("x0","x02"), "value": "%0.2f" % x0},
        { "names": ("x1",), "value": "%0.2f" % x1},
        { "names": ("f0","f02"), "value": "%0.2f" % f0},
        { "names": ("f1","f12"), "value": "%0.2f" % f1},
        { "names": ("EJ",), "value": "%0.2e" % (E*J)},
        { "names": ("EF",), "value": "%0.2e" % (E*F)},
        { "names": ("q_iz","q_iz2"), "value": "%0.2f" % q_iz},
        { "names": ("q_tr","q_tr2"), "value": "%0.2f" % q_tr},
        { "names": ("D_n","D_n2","D_n3"), "value": "%0.3f" % D_n},
        { "names": ("s_st",), "value": "%0.4f" % s_st},
        { "names": ("r_sr",), "value": "%0.3f" % r_sr},
        { "names": ("p","p2","p3"), "value": "%.1f" % (p / 10**6)},
        { "names": ("k_n",), "value": "%.2f" % k_n},
        { "names": ("s_max",), "value": "%.2f" % Ksi_s_max},
        { "names": ("p_n",), "value": config["Pipe"]["p_n"]},
        { "names": ("p_x","p_x2","p_x3","p_x4"), "value": "%.2f" % (p_x / 10**6) },
        { "names": ("ksi_N",), "value": "%.2f" % (Ksi_N/10**6)},
        { "names": ("ksi_kc_a",), "value": "%.2f" % (Ksi_kc_a/10**6)},
        { "names": ("ksi_kc_b",), "value": "%.2f" % (Ksi_kc_b/10**6)},
        { "names": ("ksi_pr_a",), "value": "%.2f" % (Ksi_pr_n_a/10**6)},
        { "names": ("ksi_pr_b",), "value": "%.2f" % (Ksi_pr_n_b/10**6)},
        { "names": ("psi",), "value": "%.2f" % psi3},
        { "names": ("u_pr_a",), "value": "%.2f" % (psi3*Usl_prochnosti/10**6)},
        { "names": ("u_pr_b",), "value": "%.2f" % (Usl_prochnosti/10**6)},
        { "names": ("L_ballast","L_ballast2","L_ballast4"), "value": "%d" % L_ballast},
        { "names": ("k_n_v",), "value": "%.2f" % k_n_v},
#        { "names": ("l_shag_18",), "value": "%.2f" % (L_ballast/1.8) },
        { "names": ("l_op",), "value": "%.2f" % L_op_shag },
        { "names": ("V_gruza",), "value": "%.2f" % V_gruza },
        { "names": ("Dlina",), "value": "%.2f" % xd[-1] },
        { "names": ("t_exp_leto","t_exp_leto1"), "value": "%.1f" % t_exp_leto },
        { "names": ("t_exp_zima","t_exp_zima1"), "value": "%.1f" % t_exp_zima },
        { "names": ("delta_t_max",), "value": "%.1f" % delta_t_max },
        { "names": ("delta_t_min",), "value": "%.1f" % delta_t_min },
]

    for l in doclist:
        for kw in l["names"]:
            try:
                worddoc.FormFields(kw).Result = l["value"]
            except:
                print(kw)
            
    #вставляем картинки
    for i in (1, 3, 4, 5):
        wordapp.Selection.GoTo(What=win32com.client.constants.wdGoToBookmark, Name="fig"+str(i))
        wordapp.Selection.InlineShapes.AddPicture(FileName=os.path.abspath(filename)+"fig"+str(i)+".png", LinkToFile=False, SaveWithDocument=True)
    
    #вставляем приложения        
    if config.has_option("Project", "appendix"):
        wordapp.Selection.GoTo(What=win32com.client.constants.wdGoToBookmark, Name="appendix")
        fnames = config["Project"]["appendix"].split(",")
        for i in fnames:
#            print(os.path.join(os.path.dirname(os.path.abspath(filename)), i.strip()))
            wordapp.Selection.InlineShapes.AddPicture(FileName=os.path.join(os.path.dirname(os.path.abspath(filename)), i.strip()), LinkToFile=False, SaveWithDocument=True)

    worddoc.Save()
    worddoc.Close() # Close the Word Document (a save-Dialog pops up)
    #wordapp.Quit() # Close the Word Application
    acad = Autocad(create_if_not_exists=True)

    pl = pline_to_acad(acad, xd, yd, 0, 0, 0.5, 5)
    pl.Color = 1
    
    pl = pline_to_acad(acad, xd, ld, 0, 0, 0.5, 5)
    pl.Color = 42    
    
    pl = pline_to_acad(acad, xd, y2d, 0, 0, 0.5, 5)
    pl.Color = 5
    
    pl = pline_to_acad(acad, xd, zd, 0, 0, 1, 1)
    pl.Color = 5  
