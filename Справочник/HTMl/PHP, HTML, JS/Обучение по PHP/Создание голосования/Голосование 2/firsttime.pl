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

#Пожалуйста удалите этот script после того, как установка закончена.

####################### Nothing more need to be modified blow this line unless you feel like to do it

require "cfg.pl";

read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});
@pairs = split(/&/, $buffer);
foreach $pair (@pairs) {
        ($name, $value) = split(/=/, $pair);
        $value =~ tr/+/ /;
        $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;
        $value =~ s/\n/ /g;

        $request{$name} = $value;


}
print "Content-type: text/html\n\n";

if ($request{'action'} eq "firsttime") {

&firsttime;

exit;

}

if ($request{'action'} eq "setup") {

&setup;

exit;

}

open (DETECT,"<$data_location/pass.dat") || &setup;
	
if ($flock eq "y") {

flock DETECT, 2; 

}

	
close (DETECT);


print "Установка закончена. Пожалуйста удалите этот скрипт. (firsttime.pl\n)";



sub setup{

print "Этот скрипт выполняется первый раз.\n";
print "Пожалуйста введите Ваш административный пароль:\n";
print "<form action=\"firsttime.pl\" method=\"post\">Пароль:<input type=\"text\" name=\"password1\"><br>Ещё раз:<input type=\"text\" name=\"password2\"><input type=\"hidden\" name=\"action\" value=\"firsttime\"><input type=\"submit\" value=\"OK\"></form>";

exit;

}

sub firsttime {
	
if ($request{'password1'} eq "" and !$request{'password1'}) {
		

   print "Пожалуйста не оставляйте пустыми поля пароля!";
		
	exit;
		
	}
	
if ($request{'password2'} eq "" and !$request{'password2'}) {
		
  
   print "Пожалуйста не оставляйте пустыми поля пароля!";
		
	exit;
		
	}

   
 if ($request{'password1'} ne "$request{'password2'}" and $request{'password1'} != $request{'password2'}){
		

   print "Вы ввели два различных пароля. Пожалуйста пробуйте снова!";
		
	exit;
		
	}

$request{'password1'}=~ tr/A-Z/a-z/; 

$password = crypt($request{'password1'}, "MM");	
    
   open (PASSWORD, ">$data_location/pass.dat") || &error("Невозможно создать файл пароля.");
    
   if ($flock eq "y") {

   flock PASSWORD, 2; 

   }

    print PASSWORD "$password";
    close (PASSWORD);
	
   open (COUNT, ">$data_location/id.txt") || &error("Невозможно создать файл данных");
    
   if ($flock eq "y") {

   flock COUNT, 2; 

   }

   print COUNT "1";
   
   close (COUNT);
   
   
   print "Установка закончена\n";
	
exit;
	
	
}


sub error{

$errors = $_[0] ;

print "Произошла ошибка,  $errors<br>\n";
print "<font color=\"red\">$!</font>\n";

exit;

}

exit;
