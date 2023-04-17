from tkinter import *
import pandas as pd
import random

class QA:
  def __init__(self, question, correctAnswer, otherAnswers):
    self.question = question
    self.corrAnsw = correctAnswer
    self.otherAnsw = otherAnswers
def question():
    # вытаскиваем вопросы из файла
    qaList =[]
    df=pd.read_excel('key.xlsx',sheet_name='Test')
    for i in range(len(df)):
        qaList.append(QA(df.loc[i,"question"],df.loc[i,"1)Correct"],[df.loc[i,"2)"],df.loc[i,"3)"],df.loc[i,"4)"]]))
    return qaList

def save(data,Name,data_result):
    # сохраняем новый лист в файл
    df1 = pd.DataFrame(data,columns=['question','correct' ,'answer','points'])
    df_result=pd.DataFrame(data_result)
    df_all=pd.concat([df1, df_result], axis=1)
    with pd.ExcelWriter('key.xlsx',  engine="openpyxl", mode='a') as writer: 
        df_all.to_excel(writer, sheet_name=Name)
def score(data,Name):
    #оценка
    clean()
    score=0
    all_score=len(data)
    for i in data:
        score+=int(i[3])
    procent=  score/(all_score/100)
    result=[]
    if procent> 85:
         result='Отлично'
         img = PhotoImage(file='otl.png')
         collor = '#008000'
    elif (procent<= 85 and procent>=70):
        result = 'Хорошо'
        img = PhotoImage(file='horosh.png')
        collor = '#00FFFF'
    elif  (procent< 70 and procent>=50):
        result = 'Удовлетворительно'
        img = PhotoImage(file='udovl.png')
        collor = '#FFFF00'
    else:
        result = 'Неудовлетворительно'
        img = PhotoImage(file='neudo.png')
        collor = '#FF0000'
    root.config(bg = collor)
    canvas = Canvas(root,width = 80, height = 80,bg = collor ,highlightbackground=collor)                              
    canvas.create_image(10,10, anchor=NW, image=img)
    res=Label(root,text=f'Баллы: {score} \n Процент: {procent}%\n Оценка: {result}', bg = collor)
    btn=Button(root,text="Главный экран",command = lambda :  first_window() , bg = '#0000FF')
    res.pack(padx=10,pady=10)
    canvas.pack()
    btn.pack(padx=10,pady=10)
    data_result = [{'score': score, 'Procent': procent,'result':result}]
    save(data,Name,data_result)
    data.clear()
    canvas.mainloop()
def logick_test(Name,answer,correct,number,p,rand_question,data,lenght):
    #логика переключения между вопросами и подсчета верности ответов
    a=answer.get()
    if (a == '1' or a == '2' or a == '3' or a== '4'):
        if (str(correct)==str(p[int(a)-1])):
            point=1
        else:
            point=0
        data.append([str(rand_question),str(correct),str(p[int(answer.get())-1]),point])
        number+=1;
        if number < lenght:
            window_question(number,qaList,Name)
        else:
            data_result=score(data,Name)
    else:
        window_question(number,qaList,Name)
def clean():
    #очистка окна
    for ele in root.winfo_children():
        ele.destroy()
def window_question(number,qaList,name):
    # окно с вопросами
    if (name):
        clean()
        lenght=len(qaList)
        p=[qaList[number].corrAnsw]+qaList[number].otherAnsw
        random.shuffle(p)
        rand_question=qaList[number].question
        correct=qaList[number].corrAnsw
        quest=Label(root,text=rand_question, bg='#00FA9A')
        var_1=Label(root,text="1)"+str(p[0]), bg='#FFFFFF',width=150 ,anchor = NW, justify=LEFT)
        var_2=Label(root,text="2)"+str(p[1]), bg='#FFFFFF',width=150,anchor = NW, justify=LEFT)
        var_3=Label(root,text="3)"+str(p[2]), bg='#FFFFFF',width=150,anchor = NW, justify=LEFT)
        var_4=Label(root,text="4)"+str(p[3]), bg='#FFFFFF',width=150,anchor = NW, justify=LEFT)
        text=Label(root,text="Напишите нномер ответа", bg='#FFFFFF')
        answer=Entry()
        btn=Button(root,text="Далее",command = lambda : logick_test(name,answer,correct,number,p,rand_question,data,lenght))
        quest.pack()
        var_1.pack(padx=5,pady=5)
        var_2.pack(padx=5,pady=5)
        var_3.pack(padx=5,pady=5)
        var_4.pack(padx=5,pady=5)
        text.pack(padx=10,pady=1)
        answer.pack()
        btn.pack(padx=10,pady=10)
    else:
        start()
   
def start():
    #начало работы теста
    clean()
    text=Label(root,text='Напишите ваше полное имя', bg='#FFFFFF')
    name=Entry()
    btn=Button(root,text="Начать тест",command = lambda :  window_question(number,qaList,name.get()), bg='#00FF00')
    text.pack(padx=10,pady=10)
    name.pack()
    btn.pack(padx=20,pady=20)
def open_base(name):
    # вывести результаты определенного человека
    if (name):
        clean()
        try :
            df_v=pd.read_excel('key.xlsx',sheet_name=name)
            text=Label(root,text=f'Баллы: {df_v.iloc[0,5]}\nПроцент: {df_v.iloc[0,6]}%\n Оценка: {df_v.iloc[0,7]}', bg='#FFFFFF')
            bk_b=Button(root,text="Назад",command=lambda : view_score())
            text.pack(padx=10,pady=10)
            bk_b.pack(padx=10,pady=10)
        except:
            view_score()
    else:
        view_score()
def table_name(data):
    Hight='300'
    # вывод в столбик
    for i in range(1,len(data)):
        Hight=int(Hight)+10
        root.geometry('300x'+str(Hight))
        text=Label(root,text=data[i], bg='#FFFFFF')
        text.pack()
def print_name():
    # вывести имена всех сдавших тест
    clean()
    xl = pd.ExcelFile('key.xlsx')
    name =xl.sheet_names 
    table_name(name)
    bk_b=Button(root,text="Назад",command=lambda : view_score(), bg='#FFFFFF')
    bk_b.pack(padx=10,pady=10)
def view_score():
    #просмотр результвтов
    root.geometry("300x350")
    clean()
    text=Label(root,text='Введите имя студента', bg='#FFFFFF')
    name_v=Entry()
    bk_s=Button(root,text="Посмотреть оценку",command=lambda : open_base(name_v.get()), bg='#FFFFFF')
    bk_v=Button(root,text="Посмотреть все доступные имена",command=lambda : print_name(), bg='#FFFFFF')
    bk_main=Button(root,text="Главный экран",command=lambda : first_window(), bg='#FFFFFF')
    text.pack(padx=10,pady=10)
    name_v.pack()
    bk_s.pack(padx=5,pady=5)
    bk_v.pack(padx=5,pady=5)
    bk_main.pack(padx=5,pady=5)
def first_window():
    #главный экран
    clean()
    root.config(bg = '#FFFFFF')
    Text=Label(root,text="Тест по истории России", font='Silkscreen 10', bg='#FFFFFF' ,fg="#DC143C")
    btn_1=Button(root,text="Старт",command = lambda : start())
    btn_1.configure(bg='green')
    btn_2=Button(root,text="Посмотреть результаты",command=lambda : view_score())
    btn_3=Button(root,text="Закрыть",command = lambda : root.destroy())
    btn_3.configure(bg='red')
    Text.pack(padx=20,pady=20)
    btn_1.pack(padx=10,pady=10)
    btn_2.pack()
    btn_3.pack(padx=10,pady=10)
############################################
#вызов нужных для запуска функций#
data=[]
qaList =question()
root=Tk()
root.title("Test")
root.geometry("300x350")
root.config(bg = '#FFFFFF')
random.shuffle(qaList)
number=0
first_window()
root.mainloop()


