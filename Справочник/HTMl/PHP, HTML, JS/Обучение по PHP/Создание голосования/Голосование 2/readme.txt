#######################################################
#             Survey V1.3	
#
# This program is distributed as freeware. We are not            	
# responsible for any damages that the program causes	
# to your system. It may be used and modified free of 
# charge, as long as the copyright notice
# in the program that give me credit remain intact.
# If you find any bugs in this program. It would be thankful
# if you can report it to us at cgifactory@cgi-factory.com.  
# However, that email address above is only for bugs reporting. 
# We will not  respond to the messages that are sent to that 
# address. If you have any trouble installing this program. 
# Please feel free to post a message on our CGI Support Forum.
# Selling this script is absolutely forbidden and illegal.
##################################################################
#
#               COPYRIGHT NOTICE:
#
#         Copyright 1999 CGI-Factory.com 
#
#      Author:  Yutung Liu
#      Web site: http://www.cgi-factory.com
#      E-Mail: cgifactory@cgi-factory.com
#      Released Date: September 27, 1999
#	
#   Survey V1.3 is protected by the copyright 
#   laws and international copyright treaties, as well as other 
#   intellectual property laws and treaties.
###################################################################


Установка:

1) Откройте файл конфигурации (cfg.pl) в простом текстовом редакторе.

И отредактируйте там все переменные в соответствии с вашими установками.

(Если путь к PERL на Вашем сервере отличается от пути по умолчанию
#!/usr/local/bin/perl 
Вы должны изменить первые строки в файлах скриптов в соответствии с Вашим путём.
(survey-admin.pl, survey.pl, result.pl, и firsttime.pl)." 

2) Откройте файлы header и footer.txt и измените их как Вам надо. (верхний и нижний клонтитулы)

3) Загрузите все скрипты в Вашу cgi-bin директорию, и установите
им атрибуты "755" (cfg.pl не требует этого) 

4) Загрузите файлы данных в  cgi-bin и установите
им атрибуты 777. (iptime.txt, ip.txt, header.txt, and footer.txt)

5) Зоздайте каталог "data," там где Вы указали его путь в строке
$data_location="/home/freyn1/cgi/survey/data"; #Путь к каталогу файлов данных.. 
в файле cfg.pl
установите его атрибут "777". Этот каталог будет использоваться для хранения файлов опросов.

6) Запустите в браузере "firsttime.pl." Это скрипт установки пароля. 
Удалите этот скрипт("firsttime.pl") после завершения установки пароля. 

7)Запустите в браузере "survey-admin.pl" для создания Ваших опросов. 

8) Done.

##########################################################################
!!Важно!!
"Все файлы надо загружать в "ASCII" режиме.
##########################################################################



