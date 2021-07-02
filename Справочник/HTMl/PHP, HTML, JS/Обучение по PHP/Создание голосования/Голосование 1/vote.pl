#!/usr/local/bin/perl
print "Content-Type: text/html\n\n";

#########################
#	Настроечки	#
#########################

$path_to_send_pl="send.pl";


open mnum, "<text/num.txt";
while (<mnum>)
{
chomp;
$num = $_;
}
close mnum;

ATNEW:
$r=rand($num+1);
$rnd=int($r);
if(($rnd < 1)||($rnd>$num+1)){goto ATNEW};

open allvotes, "<text/votes.txt";
$i=0;
while (<allvotes>)
{
$i++;
chomp;
@vote[$i] = $_;
}
close allvotes;

($question, $vara, $varb, $varc, $numa, $numb, $numc) = split(/--/, @vote[$rnd]);

print "<FORM ACTION=$path_to_send_pl?$rnd METHOD=POST>
<TABLE width=140>
    <TR>
      <TD colspan=2><font size=-2 face=arial><b><font size=-1>Голосовалка:</font></b><br>$question</font></TD>
    </TR>
    <TR>
      <TD><INPUT TYPE=RADIO VALUE=vara NAME=var><font size=-2 face=arial>$vara</font></TD>
      <TD>$numa</TD>
    </TR>
    <TR>
      <TD><INPUT TYPE=RADIO VALUE=varb NAME=var><font size=-2 face=arial>$varb</font></TD>
      <TD>$numb</TD>
    </TR>
    <TR>
      <TD><INPUT TYPE=RADIO VALUE=varc NAME=var><font size=-2 face=arial>$varc</font></TD>
      <TD>$numc</TD>
    </TR>
</TABLE>
&nbsp;&nbsp;<INPUT type=submit value=\"Клик !\">
</FORM>
";