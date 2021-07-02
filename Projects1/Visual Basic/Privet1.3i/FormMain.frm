VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FormMain 
   Caption         =   "Приветик"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1260
      Top             =   540
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   240
      Top             =   360
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------
'MS Agent
'------------------------------------------
Dim MyChar As Object 'Объект Characters
Dim LoadRequest As Object
'язык Русский
Private Const LanguageID_Ru = &H419
'-----------------------------------------------
'голос Борис
Private Const TTSModeID_Boris = "{06377F81-D48E-11D1-B17B-0020AFED142E}"
'голос Светлана
Private Const TTSModeID_Svetlana = "{06377F80-D48E-11d1-B17B-0020AFED142E}"
'-----------------------------------------------
Dim MyAgentLeft As Long
Dim MyAgentTop As Long
'------------------------------------------
Dim MyCharTTSModeID As String
'-----------------------------------------------
'MS Agent
'------------------------------------------






'--------------------------------------------------
'User Name
'--------------------------------------------------
Dim FlagMale As Boolean
Dim UserNames(1 To 5) As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'--------------------------------------------------
'User Name
'--------------------------------------------------


Dim ITimer As Integer 'время таймера 10 сек
Dim FlagStop As Boolean 'сон



Private Sub Form_Load()
'--------------------------------
'--------------------------------
'--------------------------------
On Error Resume Next
'--------------------------------
' Загружаем персонаж по умолчанию
MyAgent.Characters.Load "MyChar", "merlin.acs"
Set MyChar = MyAgent.Characters("MyChar")
'--------------------------------
If Err.Number Then
MsgBox "MS Agent не найден! Без него программа на этом компьютере работать не сможет."
Unload Me
Exit Sub
End If
'--------------------------------
On Error GoTo 0
'--------------------------------
'--------------------------------
'--------------------------------
MyChar.LanguageID = LanguageID_Ru ' Установим русский язык
'--------------------------------
'--------------------------------
'--------------------------------
On Error Resume Next
'--------------------------------
MyChar.TTSModeID = TTSModeID_Boris ' Установим движок - голос Бориса (русская версия)
'--------------------------------
'If Err.Number Then 'если ошибка
'MsgBox "Голос MS Agent не загружен!"
'End If
'--------------------------------
On Error GoTo 0
'--------------------------------
'--------------------------------
'--------------------------------
MyChar.Commands.Add "Congratulation", "&Общаться"
MyChar.Commands.Add "StopMe", "&Спать"
MyChar.Commands.Add "TimeCommand", "&Который час?"
MyChar.Commands.Add "Exit", "В&ыход"
'--------------------------------
Get_UserName 'User Name
'--------------------------------
Congratulation0
'--------------------------------
ITimer = 0
FlagStop = False
Timer1.Interval = 1500
Timer1.Enabled = True
'--------------------------------
End Sub




Private Sub Congratulation0()
'--------------------------------
MyChar.Stop
MyChar.Play "RestPose"
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
MyChar.Show ' Выводим на экран
'--------------------------------
MyChar.Play "Wave"
MyChar.Speak "Приветик, " & UserNames(NRnd(1, 5)) & "!"
MyChar.Speak "Давай пообщаемся!|Давай поговорим!|Давай поболтаем!|Давай потреплемся!"
MyChar.Play "RestPose"
'--------------------------------
End Sub




Private Sub Congratulation()
Dim IVar As Integer
'--------------------------------
IVar = NRnd(1, 2)
'--------------------------------
FlagStop = False
'--------------------------------
MyChar.Stop
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 'веселый
MyChar.Speak UserNames(NRnd(1, 5)) & ", чем занимаешься?|" & UserNames(NRnd(1, 5)) & ", что делаешь?|" & UserNames(NRnd(1, 5)) & ", как дела?|" & UserNames(NRnd(1, 5)) & ", что новенького?"
MyChar.Play "Startlistening"
MyChar.Speak "Ну, расскажи чё-нить?|Да ладно тебе - не молчи!|Какие новости?|Что нового?|Что скажешь?|Поговори со мной!"
Case 2 'грустный
MyChar.Speak UserNames(NRnd(1, 5)) & ", отстань!|" & UserNames(NRnd(1, 5)) & ", заняться нечем?|" & UserNames(NRnd(1, 5)) & ", давай работай!|" & UserNames(NRnd(1, 5)) & ", да ну тебя!|" & UserNames(NRnd(1, 5)) & ", отвали!|" & UserNames(NRnd(1, 5)) & ", не приставай!"
MyChar.Play "Uncertain"
MyChar.Speak "На фиг мне это нужно!|Лучше помолчи!|Слышать ничего не хочу!|Мне от тебя ничего не нужно!|Не очень-то и хотелось общаться!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 'веселый
    MyChar.Play "Greet"
    If FlagMale = True Then 'мужик
    MyChar.Speak UserNames(NRnd(1, 5)) & ", отлично выглядишь!|" & UserNames(NRnd(1, 5)) & ", ты парень что надо!|" & UserNames(NRnd(1, 5)) & ", ты настоящий мужик!|" & UserNames(NRnd(1, 5)) & ", я горжусь тобой!"
    MyChar.Play "Acknowledge"
    MyChar.Speak "В общем, ты молодец!|Так держать!|Рад за тебя!|Это хорошо!"
    Else
    MyChar.Speak UserNames(NRnd(1, 5)) & ", отлично выглядишь!|" & UserNames(NRnd(1, 5)) & ", ты такая красавица!|" & UserNames(NRnd(1, 5)) & ", ты самая лучшая!|" & UserNames(NRnd(1, 5)) & ", ты мне нравишься!|" & UserNames(NRnd(1, 5)) & ", я долго искал такую, как ты!"
    MyChar.Play "GestureUp"
    MyChar.Speak "Очаровашка моя!|Добрая душа!|Котёнок мой!|Цыпочка моя!|Прелесть моя!|Сокровище моё!|Звездочка!|Солнышко моё!|Зайка моя!|Радость моя!|У тебя изгибов больше, чем на железнодорожных путях!|Я буду всё делать так, как хочу, то есть по-твоему!|О, Алмаз моей души!|О, Жемчужина моего сердца!|О, Бриллиант моей судьбы!|О, Несравненная!"
    End If
Case 2 'грустный
    MyChar.Play "Decline"
    If FlagMale = True Then 'мужик
    MyChar.Speak UserNames(NRnd(1, 5)) & ", причесался бы что-ли?|" & UserNames(NRnd(1, 5)) & ", умылся бы хоть!|" & UserNames(NRnd(1, 5)) & ", что-то ты сегодня помятый!|" & UserNames(NRnd(1, 5)) & ", ну и видок у тебя!|" & UserNames(NRnd(1, 5)) & ", ты что не выспался?"
    Else
    MyChar.Speak UserNames(NRnd(1, 5)) & ", давно в зеркало смотрелась?|" & UserNames(NRnd(1, 5)) & ", сделай с собой что-нить!|" & UserNames(NRnd(1, 5)) & ", что-то ты сегодня помятая!|" & UserNames(NRnd(1, 5)) & ", ну и видок у тебя!|" & UserNames(NRnd(1, 5)) & ", сладкого объелась что-ли?|" & UserNames(NRnd(1, 5)) & ", ты что не выспалась?"
    End If
    MyChar.Play "Sad"
    MyChar.Speak "Короче, хуже некуда!|Ну и дела!|Мне грустно за тебя!|Это плохо!|Я не доволен!|Это не допустимо!|Так нельзя!|Это возмутительно!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
Select Case NRnd(1, 2) 'IVar
Case 1 'веселый
MyChar.Play "GetAttentionContinued"
MyChar.Play "GetAttentionContinued"
MyChar.Play "GetAttentionContinued"
MyChar.Speak UserNames(NRnd(1, 5)) & ", погода классная!|" & UserNames(NRnd(1, 5)) & ", на улице здорово!|" & UserNames(NRnd(1, 5)) & ", здоровский денёк выдался!|" & UserNames(NRnd(1, 5)) & ", день сегодня замечательный!"
MyChar.Play "GestureDown"
MyChar.Speak "Слушай, пойдём гулять, сколько еще ждать тебя?|Ну, пошли скорее гулять!|Ты идёшь гулять?|Погуляем, ты со мной?|Пошли гулять что-ли!"
Case 2 'грустный
MyChar.Play "Idle3_1"
MyChar.Speak UserNames(NRnd(1, 5)) & ", с тобой не интересно!|" & UserNames(NRnd(1, 5)) & ", мне скучно у тебя!|" & UserNames(NRnd(1, 5)) & ", с тобой одна тоска!|" & UserNames(NRnd(1, 5)) & ", и что я тут с тобой вожусь?"
MyChar.Play "Explain"
MyChar.Speak "Уйду я от тебя!|Я просто теряю время!|Пора мне сваливать отсюда!|Больше терпеть не могу!|Ничем не могу тебе помочь!|При чём тут я?|Я за тобой бегать не собираюсь!"
End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
'--------------------------------
'--------------------------------
MyAgentLeft = NRnd(1, 8) / 10 * Screen.Width / Screen.TwipsPerPixelX
MyAgentTop = NRnd(1, 8) / 10 * Screen.Height / Screen.TwipsPerPixelY
MyChar.MoveTo MyAgentLeft, MyAgentTop
'--------------------------------
MyChar.Play "ReadContinued"
'--------------------------------
'афоризмы
'--------------------------------
Select Case NRnd(1, 2) 'IVar

Case 1
MyChar.Speak "Если тебе нечего сказать, " & UserNames(NRnd(1, 5)) & ", то лучше помолчи.|" _
& "Если у тебя нет цели, " & UserNames(NRnd(1, 5)) & ", то не будет и результата.|" _
& "Иногда легче понять то, что тебе не нужно, " & UserNames(NRnd(1, 5)) & ", чем то, что тебе нужно.|" _
& "Обманывая других, " & UserNames(NRnd(1, 5)) & ", обманываешь прежде всего себя.|" _
& "Что может быть хуже ошибки, " & UserNames(NRnd(1, 5)) & "? - Совершить ошибку и не исправить ее.|" _
& "Чтобы принять верное решение, " & UserNames(NRnd(1, 5)) & ", нужно иметь достаточно информации.|" _
& "Делай всё, что от тебя зависит, " & UserNames(NRnd(1, 5)) & ", и ты удивишься, узнав, как много от тебя зависит!|" _
& "Поступай так, как поступил бы мудрый человек, " & UserNames(NRnd(1, 5)) & ", и ты станешь мудрым человеком.|" _
& "Плохое, " & UserNames(NRnd(1, 5)) & ", учит нас сильнее ценить хорошее.|" _
& "Когда уходят великие, " & UserNames(NRnd(1, 5)) & ", то их места занимают малые.|" _
& "Большие дела, " & UserNames(NRnd(1, 5)) & ", состоят из малых.|" _
& "Когда часто говоришь высокие слова, " & UserNames(NRnd(1, 5)) & ", то они обесцениваются.|" _
& "В каждый момент времени, " & UserNames(NRnd(1, 5)) & ", делай то, что считаешь самым важным.|" _
& "Если тебя практически всё раздражает, " & UserNames(NRnd(1, 5)) & ", значит, тебе надо отдохнуть.|" _
& "Не бойся, что люди не знают тебя, " & UserNames(NRnd(1, 5)) & ", но бойся, что ты не знаешь людей.|" _
& "Глупый человек надеется на успех, " & UserNames(NRnd(1, 5)) & ", а мудрый человек готовит свой успех.|" _
& "Глупый человек хочет, чтобы у него не было проблем, " & UserNames(NRnd(1, 5)) & ", а мудрый человек учится решать свои проблемы.|" _
& "Когда тебя застали врасплох, " & UserNames(NRnd(1, 5)) & ", не показывай этого.|" _
& "Есть вопросы, " & UserNames(NRnd(1, 5)) & ", на которые каждый человек должен ответить себе сам.|" _
& "Если ты не сделаешь свою работу, " & UserNames(NRnd(1, 5)) & ", то кто же её сделает?|" _
& "Принимать решения за другого человека легче, " & UserNames(NRnd(1, 5)) & ", чем научить его принимать решения самостоятельно.|" _
& "Мудрый человек ставит перед собой свои цели, " & UserNames(NRnd(1, 5)) & ", а глупый человек воспринимает своими - цели других людей.|" _
& "Человек не может не меняться, " & UserNames(NRnd(1, 5)) & " - он либо развивается, либо деградирует.|" _
& "Не можешь быть первым - стань лучшим, " & UserNames(NRnd(1, 5)) & ", не можешь быть лучшим - стань первым!|" _
& "Когда тяжело, " & UserNames(NRnd(1, 5)) & ", скажи себе: ''Всё будет хорошо!''"

Case 2
MyChar.Speak "Мужчина постоянно стремится подчинить себе женщину, " & UserNames(NRnd(1, 5)) & ", но когда она подчиняется, его это раздражает.|" _
& "Мужчина в отличие от женщины, " & UserNames(NRnd(1, 5)) & ", не может выполнять сразу несколько дел одновременно.|" _
& "Мужчина обдумывает решение своих прорблем молча, " & UserNames(NRnd(1, 5)) & ", а женщина - щебечет, щебечет, щебечет!|" _
& "Некоторые считают, что у них доброе сердце, хотя на самом деле, " & UserNames(NRnd(1, 5)) & ", у них лишь слабая воля.|" _
& "Женщина говорит для того, чтобы её выслушали, " & UserNames(NRnd(1, 5)) & ", а не для того, чтобы ей дали совет.|" _
& "Он сделал невозможное, " & UserNames(NRnd(1, 5)) & ", потому что не знал, что это невозможно.|" _
& "Женщины способны на всё, " & UserNames(NRnd(1, 5)) & ", а мужчины - на всё остальное.|" _
& "Мужчина, " & UserNames(NRnd(1, 5)) & ", должен быть уверен, что все важные вопросы в семье решает он.|" _
& "Все девушки такие хорошие, " & UserNames(NRnd(1, 5)) & ", откуда же тогда берутся злые жёны?|" _
& "Если мужчина даёт женщине всё, что она просит, " & UserNames(NRnd(1, 5)) & ", значит, она просит слишком мало!|" _
& "Как не редко встречается настоящая любовь, " & UserNames(NRnd(1, 5)) & ", настоящая дружба встречается ещё реже.|" _
& "Без женщин жизнь скучна, " & UserNames(NRnd(1, 5)) & ", а без дураков - невыносима!|" _
& "Лучше один раз в год родить, " & UserNames(NRnd(1, 5)) & ", чем каждый день бороду брить!|" _
& "Обояние, " & UserNames(NRnd(1, 5)) & " - это когда мужчина принимает женщину за женщину.|" _
& "Красноречие друга, " & UserNames(NRnd(1, 5)) & ", ценится выше, чем благоразумие кого-либо другого.|" _
& "Первые 20% усилий дают 80% результата, " & UserNames(NRnd(1, 5)) & ", а оставшиеся 80% усилий - всего лишь 20% результата.|" _
& "Люди, " & UserNames(NRnd(1, 5)) & ", обычно мучают своих ближних под предлогом, что желают им добра.|" _
& "Что бы люди не затевали, " & UserNames(NRnd(1, 5)) & ", их начинания будут поняты другими людьми не так, только не так, всегда не так!|" _
& "Лишь чужими глазами, " & UserNames(NRnd(1, 5)) & ", можно увидеть свои недостатки.|" _
& "Планируй своё будущее, " & UserNames(NRnd(1, 5)) & ", вместо того, чтобы беспокоиться из-за него.|" _
& "Мы, " & UserNames(NRnd(1, 5)) & ", подстраиваемся под тех, кому хотим понравиться.|" _
& "Делай то, что ты больше всего боишься, " & UserNames(NRnd(1, 5)) & ", и ты победишь страх!|" _
& "Устанавливай себе цель на каждый день, " & UserNames(NRnd(1, 5)) & ", и просматривай результаты каждый вечер.|" _
& "Быть несчастным - это привычка, быть счастливым - это тоже привычка, и выбор, " & UserNames(NRnd(1, 5)) & ", целиком зависит от тебя!|" _
& "Сердце всегда говорит правду, " & UserNames(NRnd(1, 5)) & ", а разум только то, что ты хочешь услышать."

End Select
'--------------------------------
MyChar.Play "RestPose"
'--------------------------------
End Sub



'-----------------------------------------------------
'генерация случайного числа
'-----------------------------------------------------
Private Function NRnd(Nmin As Integer, Nmax As Integer) As Integer
Randomize
NRnd = Int((Nmax - Nmin + 1) * Rnd) + Nmin 'произвольный номер
End Function
'-----------------------------------------------------
'генерация случайного числа
'-----------------------------------------------------





Private Sub MyAgent_Command(ByVal UserInput As Object)

Select Case UserInput.Name

Case "Congratulation" 'поздравление
Timer1.Enabled = False
ITimer = 0
Timer1.Interval = NRnd(8, 10) * 1000
Congratulation
Timer1.Enabled = True

Case "StopMe" 'остановить
FlagStop = True
Timer1.Enabled = False
ITimer = 0
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Play "Blink"
MyChar.Play "RestPose"
MyChar.Play "Idle3_2"

Case "TimeCommand"
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Speak UserNames(NRnd(1, 5)) & ", сейчас " & Format(Time, "h:nn")
If FlagStop = False Then
    Timer1.Interval = 1000
    Timer1.Enabled = True
Else
    MyChar.Play "RestPose"
    MyChar.Play "Blink"
    MyChar.Play "RestPose"
    MyChar.Play "Idle3_2"
End If

Case "Exit" 'закрыть
Timer1.Enabled = False
MyChar.Stop
MyChar.Play "RestPose"
MyChar.Play "Wave"
MyChar.Speak UserNames(NRnd(1, 5)) & ", пока!|" & UserNames(NRnd(1, 5)) & ", счастливо!|" & UserNames(NRnd(1, 5)) & ", до скорого!|" & UserNames(NRnd(1, 5)) & ", увидимся!|" & UserNames(NRnd(1, 5)) & ", до встречи!"
Set LoadRequest = MyChar.Hide 'выход как только Друг скрылся

End Select

End Sub


Private Sub MyAgent_RequestComplete(ByVal Request As Object)
Select Case Request
Case LoadRequest 'выход как только Друг скрылся
    MyAgent.Characters.Unload "MyChar"
    Set MyChar = Nothing
    Unload Me
End Select
End Sub



Private Sub Timer1_Timer()
ITimer = ITimer + 1
If ITimer > 5 Then 'через 5*(8...10) секунд начинаем общаться
    Timer1.Enabled = False
    ITimer = 0
    Timer1.Interval = NRnd(8, 10) * 1000
    Congratulation
    Timer1.Enabled = True
End If
End Sub


Private Sub MyAgent_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Timer1.Enabled = False
ITimer = 0 'сбрасываем таймер
Timer1.Enabled = True
End Sub








'--------------------------------------------------
'User Name
'--------------------------------------------------
Private Sub Get_UserName()
Dim strUserName As String

Select Case NRnd(1, 2) 'IVar

Case 1 'мужик
FlagMale = True

Case 2 'женщина
FlagMale = False

End Select


If FlagMale = True Then 'мужик
    UserNames(1) = "Незнакомец"
    UserNames(2) = "Незнакомец"
    UserNames(3) = "Незнакомец"
    UserNames(4) = "Незнакомец"
    UserNames(5) = "Незнакомец"
Else
    UserNames(1) = "Незнакомка"
    UserNames(2) = "Незнакомка"
    UserNames(3) = "Незнакомка"
    UserNames(4) = "Незнакомка"
    UserNames(5) = "Незнакомка"
End If

On Error GoTo ErrorHandler

strUserName = Space(255) 'Create a buffer
Call GetUserName(strUserName, 255)   'Get the username
strUserName = StripTerm(strUserName)

Select Case strUserName 'имя текущего пользователя
    
'здесь Вы обрабатываете имена всех пользователей Вашей сети
'вместо "Alexandr", "Alexandr1", "Alexandr2" и т.д. запишите
'реальные имена пользователей, например: "PoloviyAO" - мое сетевое имя

    Case "Alexandr", "Alexandr1", "Alexandr2", "PoloviyAO" 'все пользователи с именем "Александр"
    UserNames(1) = "Александр"
    UserNames(2) = "Саша"
    UserNames(3) = "Сашка"
    UserNames(4) = "Шурик"
    UserNames(5) = "Сашенька"
    FlagMale = True 'мужчина

    Case "Alexandra", "Alexandra1", "Alexandra2" 'все пользователи с именем "Александра"
    UserNames(1) = "Александра"
    UserNames(2) = "Саша"
    UserNames(3) = "Шурочка"
    UserNames(4) = "Шура"
    UserNames(5) = "Сашенька"
    FlagMale = False 'женщина
    
    Case "Alexey"
    UserNames(1) = "Алексей"
    UserNames(2) = "Лёша"
    UserNames(3) = "Лёшка"
    UserNames(4) = "Алекс"
    UserNames(5) = "Алёшенька"
    FlagMale = True

    Case "Anatoliy"
    UserNames(1) = "Анатолий"
    UserNames(2) = "Толя"
    UserNames(3) = "Толька"
    UserNames(4) = "Толян"
    UserNames(5) = "Толечка"
    FlagMale = True

    Case "Valentina"
    UserNames(1) = "Валентина"
    UserNames(2) = "Валя"
    UserNames(3) = "Валюша"
    UserNames(4) = "Валюшка"
    UserNames(5) = "Валечка"
    FlagMale = False
    
    Case "Valeriy"
    UserNames(1) = "Валерий"
    UserNames(2) = "Валера"
    UserNames(3) = "Валерка"
    UserNames(4) = "Валерчик"
    UserNames(5) = "Валерочка"
    FlagMale = True
    
    Case "Viktor"
    UserNames(1) = "Виктор"
    UserNames(2) = "Витя"
    UserNames(3) = "Витька"
    UserNames(4) = "Витёк"
    UserNames(5) = "Витечка"
    FlagMale = True

    Case "Vladimir"
    UserNames(1) = "Владимир"
    UserNames(2) = "Вова"
    UserNames(3) = "Вовка"
    UserNames(4) = "Володя"
    UserNames(5) = "Володенька"
    FlagMale = True

    Case "Vyacheslav"
    UserNames(1) = "Вячеслав"
    UserNames(2) = "Слава"
    UserNames(3) = "Славка"
    UserNames(4) = "Славик"
    UserNames(5) = "Славочка"
    FlagMale = True
    
    Case "Galina"
    UserNames(1) = "Галина"
    UserNames(2) = "Галя"
    UserNames(3) = "Галечка"
    UserNames(4) = "Галенька"
    UserNames(5) = "Галочка"
    FlagMale = False
    
    Case "Dariya"
    UserNames(1) = "Дарья"
    UserNames(2) = "Даша"
    UserNames(3) = "Дашечка"
    UserNames(4) = "Дашуля"
    UserNames(5) = "Дашенька"
    FlagMale = False

    Case "Denis"
    UserNames(1) = "Денис"
    UserNames(2) = "Деня"
    UserNames(3) = "Дениска"
    UserNames(4) = "Денисушка"
    UserNames(5) = "Денисочка"
    FlagMale = True
    
    Case "Dmitriy"
    UserNames(1) = "Дмитрий"
    UserNames(2) = "Дима"
    UserNames(3) = "Димка"
    UserNames(4) = "Диман"
    UserNames(5) = "Димочка"
    FlagMale = True

    Case "Evgeniya"
    UserNames(1) = "Евгения"
    UserNames(2) = "Женя"
    UserNames(3) = "Жень"
    UserNames(4) = "Женюшка"
    UserNames(5) = "Женечка"
    FlagMale = False

    Case "Elena"
    UserNames(1) = "Елена"
    UserNames(2) = "Лена"
    UserNames(3) = "Еленушка"
    UserNames(4) = "Ленусик"
    UserNames(5) = "Леночка"
    FlagMale = False

    Case "Ivan"
    UserNames(1) = "Иван"
    UserNames(2) = "Ваня"
    UserNames(3) = "Ванька"
    UserNames(4) = "Ванюша"
    UserNames(5) = "Ванечка"
    FlagMale = True

    Case "Igor"
    UserNames(1) = "Игорь"
    UserNames(2) = "Игорёша"
    UserNames(3) = "Игорёшка"
    UserNames(4) = "Игорёк"
    UserNames(5) = "Игорёшенька"
    FlagMale = True
    
    Case "Irina"
    UserNames(1) = "Ирина"
    UserNames(2) = "Ира"
    UserNames(3) = "Ирочка"
    UserNames(4) = "Иришка"
    UserNames(5) = "Иришенька"
    FlagMale = False
    
    Case "Kseniya"
    UserNames(1) = "Ксения"
    UserNames(2) = "Ксюша"
    UserNames(3) = "Ксенька"
    UserNames(4) = "Ксенечка"
    UserNames(5) = "Ксюшенька"
    FlagMale = False

    Case "Lubov"
    UserNames(1) = "Любовь"
    UserNames(2) = "Любаша"
    UserNames(3) = "Любушка"
    UserNames(4) = "Люба"
    UserNames(5) = "Любашенька"
    FlagMale = False

    Case "Ludmila"
    UserNames(1) = "Людмила"
    UserNames(2) = "Люда"
    UserNames(3) = "Люся"
    UserNames(4) = "Людушка"
    UserNames(5) = "Людочка"
    FlagMale = False
    
    Case "Michail"
    UserNames(1) = "Михаил"
    UserNames(2) = "Миша"
    UserNames(3) = "Мишка"
    UserNames(4) = "Миха"
    UserNames(5) = "Мишенька"
    FlagMale = True
    
    Case "Mariya"
    UserNames(1) = "Мария"
    UserNames(2) = "Маша"
    UserNames(3) = "Машечка"
    UserNames(4) = "Машуня"
    UserNames(5) = "Машенька"
    FlagMale = False
    
    Case "Nadegda"
    UserNames(1) = "Надежда"
    UserNames(2) = "Надя"
    UserNames(3) = "Надюша"
    UserNames(1) = "Наденька"
    UserNames(1) = "Надечка"
    FlagMale = False

    Case "Nataliya"
    UserNames(1) = "Наталья"
    UserNames(2) = "Наташа"
    UserNames(3) = "Наташечка"
    UserNames(4) = "Натусик"
    UserNames(5) = "Наташенька"
    FlagMale = False

    Case "Nikolay"
    UserNames(1) = "Николай"
    UserNames(2) = "Коля"
    UserNames(3) = "Колька"
    UserNames(4) = "Коль"
    UserNames(5) = "Колечка"
    FlagMale = True

    Case "Nina"
    UserNames(1) = "Нина"
    UserNames(2) = "Нинусик"
    UserNames(3) = "Нинушка"
    UserNames(4) = "Нинуля"
    UserNames(5) = "Ниночка"
    FlagMale = False

    Case "Oksana"
    UserNames(1) = "Оксана"
    UserNames(1) = "Оксан"
    UserNames(1) = "Оксанка"
    UserNames(1) = "Окси"
    UserNames(1) = "Оксаночка"
    FlagMale = False

    Case "Oleg"
    UserNames(1) = "Олег"
    UserNames(2) = "Олег"
    UserNames(3) = "Олежка"
    UserNames(4) = "Олежик"
    UserNames(5) = "Олеженька"
    FlagMale = True

    Case "Olesya"
    UserNames(1) = "Олеся"
    UserNames(2) = "Олесь"
    UserNames(3) = "Олесечка"
    UserNames(4) = "Олесенька"
    UserNames(5) = "Олесюшка"
    FlagMale = False

    Case "Olga"
    UserNames(1) = "Ольга"
    UserNames(2) = "Оля"
    UserNames(3) = "Олечка"
    UserNames(4) = "Оленька"
    UserNames(5) = "Олюшка"
    FlagMale = False

    Case "Pavel"
    UserNames(1) = "Павел"
    UserNames(2) = "Паша"
    UserNames(3) = "Пашка"
    UserNames(4) = "Палыч"
    UserNames(5) = "Павлушенька"
    FlagMale = True

    Case "Svetlana"
    UserNames(1) = "Светлана"
    UserNames(2) = "Света"
    UserNames(3) = "Светуля"
    UserNames(4) = "Светик"
    UserNames(5) = "Светочка"
    FlagMale = False

    Case "Sergey"
    UserNames(1) = "Сергей"
    UserNames(2) = "Серёжа"
    UserNames(3) = "Серёжка"
    UserNames(4) = "Серый"
    UserNames(5) = "Серёженька"
    FlagMale = True

    Case "Tatyana"
    UserNames(1) = "Татьяна"
    UserNames(2) = "Таня"
    UserNames(3) = "Танюша"
    UserNames(4) = "Танюшка"
    UserNames(5) = "Танечка"
    FlagMale = False

    Case "Yuriy"
    UserNames(1) = "Юрий"
    UserNames(2) = "Юра"
    UserNames(3) = "Юрка"
    UserNames(4) = "Юрчик"
    UserNames(5) = "Юрочка"
    FlagMale = True
    
    Case "Yuliya"
    UserNames(1) = "Юлия"
    UserNames(2) = "Юля"
    UserNames(3) = "Юлюшка"
    UserNames(4) = "Юль"
    UserNames(5) = "Юлечка"
    FlagMale = False

    Case "Jana"
    UserNames(1) = "Яна"
    UserNames(2) = "Яна"
    UserNames(3) = "Янушка"
    UserNames(4) = "Янушечка"
    UserNames(5) = "Яночка"
    FlagMale = False

'Добавьте имена своих пользователей

End Select

Exit Sub

ErrorHandler:
    
End Sub


'--------------------------------------------------
'strip the rest of the buffer
'--------------------------------------------------
Private Function StripTerm(strInput As String) As String
Dim ZeroPos As Long
StripTerm = Trim(strInput)
ZeroPos = InStr(StripTerm, Chr$(0))
If ZeroPos > 0 Then StripTerm = Left(StripTerm, ZeroPos - 1)
End Function
'--------------------------------------------------
'strip the rest of the buffer
'--------------------------------------------------
'--------------------------------------------------
'User Name
'--------------------------------------------------










