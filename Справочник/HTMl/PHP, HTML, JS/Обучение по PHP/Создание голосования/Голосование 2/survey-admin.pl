#!/usr/bin/perl

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

require "cfg.pl";

if ($ENV{'REQUEST_METHOD'} eq 'GET') {

      @pairs = split(/&/, $ENV{'QUERY_STRING'});
   }
   elsif ($ENV{'REQUEST_METHOD'} eq 'POST') {
      read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});

      @pairs = split(/&/, $buffer);
   }
   else {
      &error('request_method');
   }

   foreach $pair (@pairs) {
      ($name, $value) = split(/=/, $pair);
 
      $name =~ tr/+/ /;
      $name =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;

      $value =~ tr/+/ /;
      $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;

      $value =~ s/\n/\s/g;
    
      if ($name && $value) { $input{$name} = $value; }
   
   }

print "Content-type: text/html\n\n";


if ($input{'action'} eq "preview") { &preview }

elsif ($input{'action'} eq "add") { &add }

elsif ($input{'action'} eq "reset") { &comreset }

elsif ($input{'action'} eq "resetnow") { &resetnow }

elsif ($input{'action'} eq "view") { &view}

elsif ($input{'action'} eq "edit") { &edit}

elsif ($input{'action'} eq "pedit") { &pedit}

elsif ($input{'action'} eq "nedit") { &nedit}

elsif ($input{'action'} eq "delete") { &comerase}

elsif ($input{'action'} eq "erase") { &erase}

elsif ($input{'action'} eq "viewresults") { &viewresults}

elsif ($input{'action'} eq "radio") { &radiotags}

elsif ($input{'action'} eq "link") { &linktags}

else { &admin}

sub preview {

&vpassword;

if (!$input{'question'}) {

print "Вы не напечатали врпрос.\n";

exit;

}

if (!$input{'how'}) {

print "Вы не выбрали способ ответа.\n";

exit;

}


$count=0;

if ($input{'response'}) {

$count++;

}

if ($input{'response'}) {

$count++

}

if ($input{'response2'}) {

$count++

}

if ($input{'response3'}) {

$count++


}

if ($input{'response4'}) {

$count++


}

if ($input{'response5'}) {

$count++


}

if ($count<2) {
	
print "Надо напечатать по меньшей мере один ответ.\n";

exit;	

}

print "Следующие вещи будут добавлены<br>\n";

print "Вопрос: $input{'question'}<br>\n";

if ($input{'response'}) {


print "Ответ: $input{'response'}<br>\n";


}

if ($input{'response2'}) {


print "Ответ: $input{'response2'}<br>\n";


}

if ($input{'response3'}) {


print "Ответ: $input{'response3'}<br>\n";


}

if ($input{'response4'}) {


print "Ответ: $input{'response4'}<br>\n";


}

if ($input{'response5'}) {


print "Ответ: $input{'response5'}<br>\n";


}

print "<form action=\"survey-admin.pl\" method=\"post\">";
print "<input type=\"hidden\" name=\"question\" value=\"$input{'question'}\">";

if ($input{'response'}) {

print "<input type=\"hidden\" name=\"response\" value=\"$input{'response'}\">";

}

if ($input{'response2'}) {

print "<input type=\"hidden\" name=\"response2\" value=\"$input{'response2'}\">";

}

if ($input{'response3'}) {

print "<input type=\"hidden\" name=\"response3\" value=\"$input{'response3'}\">";

}

if ($input{'response4'}) {

print "<input type=\"hidden\" name=\"response4\" value=\"$input{'response4'}\">";

}

if ($input{'response5'}) {

print "<input type=\"hidden\" name=\"response5\" value=\"$input{'response5'}\">";

}
print "<input type=\"hidden\" name=\"how\" value=\"$input{'how'}\">";
print "<input type=\"hidden\" name=\"password\" value=\"$input{'password'}\">";
print "<input type=\"hidden\" name=\"action\" value=\"add\">";
print "<input type=\"submit\" value=\"Add\">";
print "</form>";

}

sub add{

&vpassword;

open (id, "<$data_location/id.txt") or &error("Невозможно отврыть id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
$id=<id>;

close(id);

if ($id=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }

open (data, ">$data_location/$id.txt") or &error("Невозможно создать файл данных");

if ($flock eq "y") {

    flock data, 2; 

     }

print data "0\n";

print data "$input{'question'}\n";

if ($input{'response'}) {

print data "$input{'response'}-x%x-0\n";


}

if ($input{'response2'}) {

print data "$input{'response2'}-x%x-0\n";


}

if ($input{'response3'}) {

print data "$input{'response3'}-x%x-0\n";


}

if ($input{'response4'}) {

print data "$input{'response4'}-x%x-0\n";


}

if ($input{'response5'}) {

print data "$input{'response5'}-x%x-0\n";


}


close (data);

$id++;

open (id, ">$data_location/id.txt") or &error("Невозможно прочитать id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
print id "$id";

close(id);

$id--;

print "Новый опрос добавлен.<br>\n";

if ($input{'how'} eq "link") {

print "Нажмите сюда <a href=\"survey-admin.pl?action=link&id=$id&password=$input{'password'}\">сюда </a> что бы получить HTML код для вставки в Ваши страницы<br>\n";
print "Вы можете получить это код в любое время воспользовавшись функцие \"посмотреть\" на административной странице.<br>\n";

}

else {
	
print "Нажмите сюда <a href=\"survey-admin.pl?action=radio&id=$id&password=$input{'password'}\">сюда </a> что бы получить HTML код для вставки в Ваши страницы<br>\n";
print "Вы можете получить это код в любое время воспользовавшись функцие \"посмотреть\" на административной странице.<br>\n";

}	

exit;

}

sub view {

&vpassword;

open (id, "<$data_location/id.txt") or &error("Невозможно открыть id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
$id=<id>;

close(id);

$id--;

if ($id==0) {	
	
print "Вы не создали никакие опросы, пока...";

exit;

}

$x=1;

Opening:for ($x; $x<=$id; $x++) {

open (data, "<$data_location/$x.txt") or next Opening;

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "<b>ID:</b> $x<br>\n";
print "@data[1]<br>\n";
print "<a href=\"survey-admin.pl?action=viewresults&id=$x&password=$input{'password'}\">Результаты</a>---<a href=\"survey-admin.pl?action=edit&id=$x&password=$input{'password'}\">Изменить</a>---<a href=\"survey-admin.pl?action=reset&id=$x&password=$input{'password'}\">Обнулить </a>---<a href=\"survey-admin.pl?action=delete&id=$x&password=$input{'password'}\">Удалить этот опрос</a>---<a href=\"survey-admin.pl?action=radio&id=$x&password=$input{'password'}\">Получить HTML код(переключатели)</a>---<a href=\"survey-admin.pl?action=link&id=$x&password=$input{'password'}\">Получить HTML код (ссылки)</a>\n";  
print "<br><br>\n";

}

print "-Конец результатов поиска\n";

exit;

}


sub comerase {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты! Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "Нечего удалить!";

exit;
	
}	

open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "Вы уверены что хотите удалить этот опрос?<br><br>";

print "<b>Вопрос:</b><br>@data[1]\n<br><br>";


print "<a href=\"survey-admin.pl?action=erase&id=$input{'id'}&password=$input{'password'}\">Да, удаляем.</a>\n";


exit;

}	

sub erase {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "Нечего удалять!";

exit;
	
}	

unlink ("$input{'id'}.txt");

print "Опроса нет. $input{'id'} Был удален.\n";

exit;

}

sub viewresults {
	
if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }


if (!$input{'id'}) {
	
print "Нечего посмотреть.\n";

exit;	
	
}		
	
open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "<b>Вопрос:</b> @data[1]<br>";

$base=@data[0];

splice(@data, 0,2);
print "Результаты Голосования:<br>\n";

if ($base==0) {
	
print "Отчёт о голосовании пуст.\n";

exit;	

}

foreach $data (@data) {	
	
@datax= split(/-x%x-/, $data);

$cal=@datax[1]/$base;
$percent=100*$cal;
$percent=int($percent);

print "<b>Ответ:</b> @datax[0], @datax[1] голосов <b>($percent%)</b><br>\n";

}

}

sub comreset {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "Нечего обнулять!";

exit;
	
}	

open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "Вы уверены что хотите обнулить этот опрос?<br><br>";

print "<b>Вопрос:</b><br>@data[1]\n<br><br>";


print "<a href=\"survey-admin.pl?action=resetnow&id=$input{'id'}&password=$input{'password'}\">Да, обнуляем.</a>\n";


exit;
	
}	

sub resetnow {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }
        
if (!$input{'id'}) {
	
print "Нечего обнулять!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);


chomp (@data[2]);
chomp (@data[3]);
chomp (@data[4]);
chomp (@data[5]);
chomp (@data[6]);


@data[2]=~ s/-x%x-.*/-x%x-0\n/g;
@data[3]=~ s/-x%x-.*/-x%x-0\n/g;
@data[4]=~ s/-x%x-.*/-x%x-0\n/g;
@data[5]=~ s/-x%x-.*/-x%x-0\n/g;
@data[6]=~ s/-x%x-.*/-x%x-0\n/g;

open (data, ">$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
print data "0\n";
print data @data[1];
print data @data[2];
print data @data[3];
print data @data[4];
print data @data[5];
print data @data[6];

close(data);

print "Опроса нет. $input{'id'} был обнулён.\n";

exit;	
	
}

sub edit {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "Нечего изменить!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

chomp (@data[1]);
chomp (@data[2]);
chomp (@data[3]);
chomp (@data[4]);
chomp (@data[5]);
chomp (@data[6]);

@data[2]=~ s/-x%x-.*//g;
@data[3]=~ s/-x%x-.*//g;
@data[4]=~ s/-x%x-.*//g;
@data[5]=~ s/-x%x-.*//g;
@data[6]=~ s/-x%x-.*//g;


print "<table border=\"0\">\n";
print "<form action=\"survey-admin.pl\" method=\"post\">\n";
print "<tr><td bgcolor=\"#9F590B\"><font color=\"White\">Изменение опроса</font><br></td><td></td></tr>\n";
print "<tr><td valign=\"top\"><b>Вопрос:</b></td><td><input type=\"text\" name=\"question\" size=\"80\" value=\"@data[1]\"></td></tr>\n";
print "<tr><td><b>Ответ 1:</b></td><td><input type=\"text\" name=\"response\" size=\"30\" value=\"@data[2]\"></td></tr>\n";
print "<tr><td><b>Ответ  2:</b></td><td><input type=\"text\" name=\"response2\" size=\"30\" value=\"@data[3]\"></td></tr>\n";
print "<tr><td><b>Ответ  3:</b></td><td><input type=\"text\" name=\"response3\" size=\"30\" value=\"@data[4]\"></td></tr>\n";
print "<tr><td><b>Ответ  4:</b></td><td><input type=\"text\" name=\"response4\" size=\"30\" value=\"@data[5]\"></td></tr>\n";
print "<tr><td><b>Ответ  5:</b></td><td><input type=\"text\" name=\"response5\" size=\"30\" value=\"@data[6]\"></td></tr>\n";
print "<tr><td><b>Пароль:</b></td><td><input type=\"text\" name=\"password\"></td></tr>\n";
print "<tr><td></td><td><input type=\"hidden\" name=\"id\" value=\"$input{'id'}\"><input type=\"hidden\" name=\"action\" value=\"nedit\">\n";
print "<input type=\"submit\" value=\"Изменить\">\n";
print "</form></table>\n";	
print "<i>Внимание!!! При изменении опроса результаты будут обнулены.</i>";

exit;	
	
}

sub nedit {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "Нечего изменить!";

exit;
	
}	

	
open (data, ">$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
print data "0\n";
print data "$input{'question'}\n";

if ($input{'response'}) {
print data "$input{'response'}-x%x-0\n";
}

if ($input{'response2'}) {
print data "$input{'response2'}-x%x-0\n";
}

if ($input{'response3'}) {
print data "$input{'response3'}-x%x-0\n";
}

if ($input{'response4'}) {
print data "$input{'response4'}-x%x-0\n";
}

if ($input{'response5'}) {
print data "$input{'response5'}-x%x-0\n";
}

close(data);
	
print "Сообщений нет. $input{'id'} был изменён";

exit;

}

sub radiotags {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "Нечего показать!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

chomp (@data[1]);
chomp (@data[2]);
chomp (@data[3]);
chomp (@data[4]);
chomp (@data[5]);
chomp (@data[6]);


@data[2]=~ s/-x%x-.*//g;
@data[3]=~ s/-x%x-.*//g;
@data[4]=~ s/-x%x-.*//g;
@data[5]=~ s/-x%x-.*//g;
@data[6]=~ s/-x%x-.*//g;
	
	
print "Вставьте следующий код в Ваши страницы,<br>\n";
print "&lt;!---start here---&gt;<br>\n";
print "@data[1]&lt;br&gt;<br><br>\n";

print "&lt;form action=\"$cgi_path\" method=\"post\"&gt;";
print "&lt;input type=\"hidden\" name=\"id\" value=\"$input{'id'}\"&gt;<br>\n";
if (@data[2]) {
	
print "&lt;input type=\"radio\" name=\"response\" value=\"1\" checked&gt;@data[2]&lt;br&gt;<br>\n";

}

if (@data[3]) { 

print "&lt;input type=\"radio\" name=\"response\" value=\"2\"&gt;@data[3]&lt;br&gt;<br>\n";

}

if (@data[4]) {
	
print "&lt;input type=\"radio\" name=\"response\" value=\"3\"&gt;@data[4]&lt;br&gt;<br>\n";

}

if (@data[5]) {

print "&lt;input type=\"radio\" name=\"response\" value=\"4\"&gt;@data[5]&lt;br&gt;<br>\n";

}

if (@data[6]) {

print "&lt;input type=\"radio\" name=\"response\" value=\"5\"&gt;@data[6]&lt;br&gt;<br>\n";

}

print "&lt;br&gt;&lt;input type=\"submit\" value=\"Vote!\"&gt;\n";
print "&lt;/form&gt;<br>\n";
print "&lt;a href=\"$result?id=$input{'id'}\"&gt;Посмотреть результаты не голосуя&lt;/a&gt;&lt;br&gt;<br>\n";
print "&lt;!---end here---&gt;<br>\n";


exit;

}

sub linktags {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "Предупреждение Защиты!  Действие отменено.<br>\n";
        	print "Пожалуйста не используйте сверхъестественные символы\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "Нечего показать!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("Невозможно открыть файл данных.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

chomp (@data[1]);
chomp (@data[2]);
chomp (@data[3]);
chomp (@data[4]);
chomp (@data[5]);
chomp (@data[6]);

@data[2]=~ s/-x%x-.*//g;
@data[3]=~ s/-x%x-.*//g;
@data[4]=~ s/-x%x-.*//g;
@data[5]=~ s/-x%x-.*//g;
@data[6]=~ s/-x%x-.*//g;

	

print "Вставьте следующий код в Ваши страницы,<br>\n";

print "&lt;!---start here---&gt;<br>\n";
print "@data[1]&lt;br&gt;<br><br>\n";

if (@data[2]) {
	
print "&lt;a href=\"$cgi_path?id=$input{'id'}&response=1\"&gt;@data[2]&lt;/a&gt;&lt;br&gt;<br>\n";

}

if (@data[3]) { 

print "&lt;a href=\"$cgi_path?id=$input{'id'}&response=2\"&gt;@data[3]&lt;/a&gt;&lt;br&gt;<br>\n";

}

if (@data[4]) {
	
print "&lt;a href=\"$cgi_path?id=$input{'id'}&response=3\"&gt;@data[4]&lt;/a&gt;&lt;br&gt;<br>\n";

}

if (@data[5]) {

print "&lt;a href=\"$cgi_path?id=$input{'id'}&response=4\"&gt;@data[5]&lt;/a&gt;&lt;br&gt;<br>\n";

}

if (@data[6]) {

print "&lt;a href=\"$cgi_path?id=$input{'id'}&response=5\"&gt;@data[6]&lt;/a&gt;&lt;br&gt;<br>\n";

}
print "<br>&lt;br&gt;&lt;a href=\"$result?id=$input{'id'}\"&gt;Посмотреть результаты не голосуя&lt;/a&gt;&lt;br&gt;<br>\n";

print "&lt;!---end here---&gt;<br>\n";	
exit;

}							
#####error
sub error {
print "Произошла ошибка. <br>  $_[0]<br>\n";	
print "$!\n";	
exit;	
}
sub admin {
print <<EOF;
<html>
<head>
<title>
Наши опросы.
</title>
</head>
<body bgcolor="#ffffff">
<center>
<font face="Arial">
<b><font size="+3">Survey V1.3 Admin</font></b><font size="-1">&nbsp;&nbsp;<a href="http://www.cgi-factory.com/">from CGI-Factory.com!</a></font>
<br>
<b><font size="">Перевод</font></b><font size="-1">&nbsp;&nbsp;<a href="http://www.freyn.agava.ru/">ВизиТ</a></font>
<br>
<table border="0">
<form action="survey-admin.pl" method="post">
<tr><td bgcolor="#9F590B"><font color="White">Создать новый опрос?</font><br></td><td></td></tr>
<tr><td valign="top"><b>Вопрос:</b></td><td><input type="text" name="question" size="80"></td></tr>
<tr><td><b>Ответ 1:</b></td><td><input type="text" name="response" size="30"></td></tr>
<tr><td><b>Ответ 2:</b></td><td><input type="text" name="response2" size="30"></td></tr>
<tr><td><b>Ответ 3:</b></td><td><input type="text" name="response3" size="30"></td></tr>
<tr><td><b>Ответ 4:</b></td><td><input type="text" name="response4" size="30"></td></tr>
<tr><td><b>Ответ 5:</b></td><td><input type="text" name="response5" size="30"></td></tr>
<tr><td><b>Как будет ваш посетитель голосовать?</b></td><td>Нажатием по текстовой ссылке <input type="radio" name="how" value="link" checked><br>  Переключателями<input type="radio" name="how" value="radio"> </td></tr>
<tr><td><b>Административный пароль:</b></td><td><input type="text" name="password"></td></tr>
<tr><td></td><td><input type="hidden" name="action" value="preview">
<input type="submit" value="Создать">
</form><br><br></td></tr><tr bgcolor="#9F590B"><td><font color="White">Посмотреть все опросы?</font></td></tr>
<tr><td><form action="survey-admin.pl" method="POST">
<input type="hidden" name="action" value="view">
<td><b>Административный пароль:</b><input type="text" name="password">
<input type="Submit" value="Да"></td>
</form></td></tr></table>
</font></center>
<font size="-1"><center>&copy;1999 CGI-Factory.com</center></font>
</body>
</html>

EOF
exit;
}

sub vpassword{
open (PASS, "<$data_location/pass.dat") || &error("Невозможно открыть файл пароля"); 
if ($flock eq "y") {
flock PASS, 2; 
}
$pass = <PASS>;
close(PASS);
$input{'password'}=~ tr/A-Z/a-z/;
$pass2 = crypt($input{'password'}, "MM");
unless ($pass eq "$pass2") {
	
$timenow=localtime();
	
	print "Неверный пвроль. Вернитесь назад и попробуйте ещё раз.<br>";
		print "Пароль, который Вы ввели,  неправилен.<br>";
        print "Следующая информация была послана администрации этого сайта<br>";
        print "Ваша Информация: <ul>$ENV{'REMOTE_HOST'}</ul>";
        print "<ul>Пароль: $form{'password'}</ul>";
        print "<ul>Время: $timenow</ul>";
        print "</body></html>";
 
 if ($alert eq "y") {       
        
        open (MAIL, "|$mailprog -t") or &error("Unable to open the mail program");
        print MAIL "To: $yourmail\n";
        print MAIL "From: $yourmail\n";
        print MAIL "Subject: bad password\n";
        print MAIL "Кто то\n";
        print MAIL "Ввёл неправильный пароль.\n";
        print MAIL "Имеется информация:\n\n";
        print MAIL "$ENV{'REMOTE_ADDR'}\n";
        print MAIL "Пароль: $input{'password'}\n";
        print MAIL "$timenow\n";
        
        close(MAIL);
        
        exit;
        	
        }      
   
exit;
	
}

}	
