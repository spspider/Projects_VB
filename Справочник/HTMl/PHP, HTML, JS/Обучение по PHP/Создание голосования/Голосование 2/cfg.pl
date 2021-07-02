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

$result=golosovanie/result.pl";
#WWW путь к файлу result.pl.

$cgi_path=golosovanie/survey.pl";
#WWW путь к файлу survey.pl.

$bar="golosovanie/bar.gif";
#WWW путь к bar.gif.

$data_location="data";
#Путь к каталогу файлов данных.

$ip_logging="1";
#Регистрация IP посетителя что бы предотвратить многократные голосования. 1= да, 0 = нет

$log_clean="60";
#Период модификации(в секундах). 1800=пол часа. 3600=час. 86400=24 часа. И т.д.

$mailprog="/usr/sbin/sendmail -t";
#Путь к sendmail.

$yourmail="xsting\@rambler.ru";
#Ваш email адрес. Обратный слеш -\ после знака @ нужен

$alert="y";
#Скрипт пришлёт Вам предупреждение если кто нибудь введёт неправильный пароль. (n=нет, y=да)
$flock="y";
#Блокировка файлов (n=нет, y=да).