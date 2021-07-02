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

print "�� �� ���������� ������.\n";

exit;

}

if (!$input{'how'}) {

print "�� �� ������� ������ ������.\n";

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
	
print "���� ���������� �� ������� ���� ���� �����.\n";

exit;	

}

print "��������� ���� ����� ���������<br>\n";

print "������: $input{'question'}<br>\n";

if ($input{'response'}) {


print "�����: $input{'response'}<br>\n";


}

if ($input{'response2'}) {


print "�����: $input{'response2'}<br>\n";


}

if ($input{'response3'}) {


print "�����: $input{'response3'}<br>\n";


}

if ($input{'response4'}) {


print "�����: $input{'response4'}<br>\n";


}

if ($input{'response5'}) {


print "�����: $input{'response5'}<br>\n";


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

open (id, "<$data_location/id.txt") or &error("���������� ������� id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
$id=<id>;

close(id);

if ($id=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }

open (data, ">$data_location/$id.txt") or &error("���������� ������� ���� ������");

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

open (id, ">$data_location/id.txt") or &error("���������� ��������� id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
print id "$id";

close(id);

$id--;

print "����� ����� ��������.<br>\n";

if ($input{'how'} eq "link") {

print "������� ���� <a href=\"survey-admin.pl?action=link&id=$id&password=$input{'password'}\">���� </a> ��� �� �������� HTML ��� ��� ������� � ���� ��������<br>\n";
print "�� ������ �������� ��� ��� � ����� ����� ���������������� ������� \"����������\" �� ���������������� ��������.<br>\n";

}

else {
	
print "������� ���� <a href=\"survey-admin.pl?action=radio&id=$id&password=$input{'password'}\">���� </a> ��� �� �������� HTML ��� ��� ������� � ���� ��������<br>\n";
print "�� ������ �������� ��� ��� � ����� ����� ���������������� ������� \"����������\" �� ���������������� ��������.<br>\n";

}	

exit;

}

sub view {

&vpassword;

open (id, "<$data_location/id.txt") or &error("���������� ������� id.txt");

if ($flock eq "y") {

    flock id, 2; 

     }
     
$id=<id>;

close(id);

$id--;

if ($id==0) {	
	
print "�� �� ������� ������� ������, ����...";

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
print "<a href=\"survey-admin.pl?action=viewresults&id=$x&password=$input{'password'}\">����������</a>---<a href=\"survey-admin.pl?action=edit&id=$x&password=$input{'password'}\">��������</a>---<a href=\"survey-admin.pl?action=reset&id=$x&password=$input{'password'}\">�������� </a>---<a href=\"survey-admin.pl?action=delete&id=$x&password=$input{'password'}\">������� ���� �����</a>---<a href=\"survey-admin.pl?action=radio&id=$x&password=$input{'password'}\">�������� HTML ���(�������������)</a>---<a href=\"survey-admin.pl?action=link&id=$x&password=$input{'password'}\">�������� HTML ��� (������)</a>\n";  
print "<br><br>\n";

}

print "-����� ����������� ������\n";

exit;

}


sub comerase {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������! �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "������ �������!";

exit;
	
}	

open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "�� ������� ��� ������ ������� ���� �����?<br><br>";

print "<b>������:</b><br>@data[1]\n<br><br>";


print "<a href=\"survey-admin.pl?action=erase&id=$input{'id'}&password=$input{'password'}\">��, �������.</a>\n";


exit;

}	

sub erase {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "������ �������!";

exit;
	
}	

unlink ("$input{'id'}.txt");

print "������ ���. $input{'id'} ��� ������.\n";

exit;

}

sub viewresults {
	
if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }


if (!$input{'id'}) {
	
print "������ ����������.\n";

exit;	
	
}		
	
open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "<b>������:</b> @data[1]<br>";

$base=@data[0];

splice(@data, 0,2);
print "���������� �����������:<br>\n";

if ($base==0) {
	
print "����� � ����������� ����.\n";

exit;	

}

foreach $data (@data) {	
	
@datax= split(/-x%x-/, $data);

$cal=@datax[1]/$base;
$percent=100*$cal;
$percent=int($percent);

print "<b>�����:</b> @datax[0], @datax[1] ������� <b>($percent%)</b><br>\n";

}

}

sub comreset {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	

open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

if ($flock eq "y") {

flock data, 2; 

}
@data=<data>;

close(data);

print "�� ������� ��� ������ �������� ���� �����?<br><br>";

print "<b>������:</b><br>@data[1]\n<br><br>";


print "<a href=\"survey-admin.pl?action=resetnow&id=$input{'id'}&password=$input{'password'}\">��, ��������.</a>\n";


exit;
	
}	

sub resetnow {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }
        
if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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

open (data, ">$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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

print "������ ���. $input{'id'} ��� ������.\n";

exit;	
	
}

sub edit {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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
print "<tr><td bgcolor=\"#9F590B\"><font color=\"White\">��������� ������</font><br></td><td></td></tr>\n";
print "<tr><td valign=\"top\"><b>������:</b></td><td><input type=\"text\" name=\"question\" size=\"80\" value=\"@data[1]\"></td></tr>\n";
print "<tr><td><b>����� 1:</b></td><td><input type=\"text\" name=\"response\" size=\"30\" value=\"@data[2]\"></td></tr>\n";
print "<tr><td><b>�����  2:</b></td><td><input type=\"text\" name=\"response2\" size=\"30\" value=\"@data[3]\"></td></tr>\n";
print "<tr><td><b>�����  3:</b></td><td><input type=\"text\" name=\"response3\" size=\"30\" value=\"@data[4]\"></td></tr>\n";
print "<tr><td><b>�����  4:</b></td><td><input type=\"text\" name=\"response4\" size=\"30\" value=\"@data[5]\"></td></tr>\n";
print "<tr><td><b>�����  5:</b></td><td><input type=\"text\" name=\"response5\" size=\"30\" value=\"@data[6]\"></td></tr>\n";
print "<tr><td><b>������:</b></td><td><input type=\"text\" name=\"password\"></td></tr>\n";
print "<tr><td></td><td><input type=\"hidden\" name=\"id\" value=\"$input{'id'}\"><input type=\"hidden\" name=\"action\" value=\"nedit\">\n";
print "<input type=\"submit\" value=\"��������\">\n";
print "</form></table>\n";	
print "<i>��������!!! ��� ��������� ������ ���������� ����� ��������.</i>";

exit;	
	
}

sub nedit {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }

if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	

	
open (data, ">$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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
	
print "��������� ���. $input{'id'} ��� ������";

exit;

}

sub radiotags {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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
	
	
print "�������� ��������� ��� � ���� ��������,<br>\n";
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
print "&lt;a href=\"$result?id=$input{'id'}\"&gt;���������� ���������� �� �������&lt;/a&gt;&lt;br&gt;<br>\n";
print "&lt;!---end here---&gt;<br>\n";


exit;

}

sub linktags {

&vpassword;

if ($input{'id'}=~ tr/;<>*|`&$!#()[]{}:'"//) {
        	
            print "�������������� ������!  �������� ��������.<br>\n";
        	print "���������� �� ����������� ������������������ �������\n";
        	
        	exit;
        	
        }
	
if (!$input{'id'}) {
	
print "������ ��������!";

exit;
	
}	
	
open (data, "<$data_location/$input{'id'}.txt") or &error("���������� ������� ���� ������.");

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

	

print "�������� ��������� ��� � ���� ��������,<br>\n";

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
print "<br>&lt;br&gt;&lt;a href=\"$result?id=$input{'id'}\"&gt;���������� ���������� �� �������&lt;/a&gt;&lt;br&gt;<br>\n";

print "&lt;!---end here---&gt;<br>\n";	
exit;

}							
#####error
sub error {
print "��������� ������. <br>  $_[0]<br>\n";	
print "$!\n";	
exit;	
}
sub admin {
print <<EOF;
<html>
<head>
<title>
���� ������.
</title>
</head>
<body bgcolor="#ffffff">
<center>
<font face="Arial">
<b><font size="+3">Survey V1.3 Admin</font></b><font size="-1">&nbsp;&nbsp;<a href="http://www.cgi-factory.com/">from CGI-Factory.com!</a></font>
<br>
<b><font size="">�������</font></b><font size="-1">&nbsp;&nbsp;<a href="http://www.freyn.agava.ru/">�����</a></font>
<br>
<table border="0">
<form action="survey-admin.pl" method="post">
<tr><td bgcolor="#9F590B"><font color="White">������� ����� �����?</font><br></td><td></td></tr>
<tr><td valign="top"><b>������:</b></td><td><input type="text" name="question" size="80"></td></tr>
<tr><td><b>����� 1:</b></td><td><input type="text" name="response" size="30"></td></tr>
<tr><td><b>����� 2:</b></td><td><input type="text" name="response2" size="30"></td></tr>
<tr><td><b>����� 3:</b></td><td><input type="text" name="response3" size="30"></td></tr>
<tr><td><b>����� 4:</b></td><td><input type="text" name="response4" size="30"></td></tr>
<tr><td><b>����� 5:</b></td><td><input type="text" name="response5" size="30"></td></tr>
<tr><td><b>��� ����� ��� ���������� ����������?</b></td><td>�������� �� ��������� ������ <input type="radio" name="how" value="link" checked><br>  ���������������<input type="radio" name="how" value="radio"> </td></tr>
<tr><td><b>���������������� ������:</b></td><td><input type="text" name="password"></td></tr>
<tr><td></td><td><input type="hidden" name="action" value="preview">
<input type="submit" value="�������">
</form><br><br></td></tr><tr bgcolor="#9F590B"><td><font color="White">���������� ��� ������?</font></td></tr>
<tr><td><form action="survey-admin.pl" method="POST">
<input type="hidden" name="action" value="view">
<td><b>���������������� ������:</b><input type="text" name="password">
<input type="Submit" value="��"></td>
</form></td></tr></table>
</font></center>
<font size="-1"><center>&copy;1999 CGI-Factory.com</center></font>
</body>
</html>

EOF
exit;
}

sub vpassword{
open (PASS, "<$data_location/pass.dat") || &error("���������� ������� ���� ������"); 
if ($flock eq "y") {
flock PASS, 2; 
}
$pass = <PASS>;
close(PASS);
$input{'password'}=~ tr/A-Z/a-z/;
$pass2 = crypt($input{'password'}, "MM");
unless ($pass eq "$pass2") {
	
$timenow=localtime();
	
	print "�������� ������. ��������� ����� � ���������� ��� ���.<br>";
		print "������, ������� �� �����,  ����������.<br>";
        print "��������� ���������� ���� ������� ������������� ����� �����<br>";
        print "���� ����������: <ul>$ENV{'REMOTE_HOST'}</ul>";
        print "<ul>������: $form{'password'}</ul>";
        print "<ul>�����: $timenow</ul>";
        print "</body></html>";
 
 if ($alert eq "y") {       
        
        open (MAIL, "|$mailprog -t") or &error("Unable to open the mail program");
        print MAIL "To: $yourmail\n";
        print MAIL "From: $yourmail\n";
        print MAIL "Subject: bad password\n";
        print MAIL "��� ��\n";
        print MAIL "��� ������������ ������.\n";
        print MAIL "������� ����������:\n\n";
        print MAIL "$ENV{'REMOTE_ADDR'}\n";
        print MAIL "������: $input{'password'}\n";
        print MAIL "$timenow\n";
        
        close(MAIL);
        
        exit;
        	
        }      
   
exit;
	
}

}	
