#!/usr/local/bin/perl
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

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
print "Предупреждение Защиты!  Действие отменено.<br>\n";
print "Пожалуйста не используйте сверхъестественные символы\n";
exit;
}
open (count, "<$data_location/id.txt") or &error("It is impossible to open id.txt.");
if ($flock eq "y") {
flock count, 2; 
}
$count=<count>;
close(count);
if (!$input{'id'} or $input{'id'} >= $count) {
&header;
print "Error: Polling-$input{'id'} not found.";
&footer;
exit;
}	

open (data, "<$data_location/$input{'id'}.txt") or &error("It is impossible to open the data file.");
if ($flock eq "y") {
flock data, 2; 
}
@data=<data>;
close(data);

&header;
print "<h2>@data[1]</h2>";
$base=@data[0];
splice(@data, 0,2);
foreach $data (@data) {	
@datax= split(/-x%x-/, $data);
if ($base==0) {
$cal=0;
}
else {
$cal=@datax[1]/$base;
}
$percent=sprintf("%.2f", 100*$cal);
print "<b>@datax[0]</b> <font size=\"-1\">(@datax[1] голосов)</font> <img src=\"$bar\" width=\"$percent\" height=\"10\"> <font size=\"-1\"><b>$percent%</b></font><br>\n";	
}
print "(Всего: $base)";
print " <br><br> <br> <br> <center>  <a href=\"http://www.webmastersonline.hpage.net\">Polling from WebmastersOnline</a> </center>  ";
&footer;
######################input 
sub header { 
open (header, "<header.txt") or &error(" It is impossible to open file  header.txt "); 
if ($flock eq "y") { 
flock header, 2;
} 
@header=<header>; 
close(header);
print @header;
}
sub footer { 
open (footer, "<footer.txt") or &error("It is impossible to open footer.txt"); 
if ($flock eq "y") { 
flock footer, 2;
} 
@footer=<footer>; 
close(footer);
print @footer;
}
exit;
#####error
sub error {
print "Error. <br>  $_[0]<br>\n";	
print "$!\n";	
exit;	
}
