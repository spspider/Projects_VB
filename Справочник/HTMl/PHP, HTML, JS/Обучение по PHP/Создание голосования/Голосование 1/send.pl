#!/usr/local/bin/perl
print "Content-Type: text/html\n\n";

#########################
#	Настроечки	#
#########################

$go_url="test.htm>";

read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});
@pairs = split(/&/, $buffer);

foreach $pair (@pairs)
{
($var, $value) = split(/=/, $pair);
if($var eq "var"){$avar = $value};
}

open votes, "<text/votes.txt";
$i=0;
while (<votes>)
{
$i++;
@vote[$i] = $_;
}
close votes;

=head
open numfile, "<text/num.txt";
$x=0;
while (<numfile>)
{
$x++;
chomp;
if($x==1){$anum=$_}
}
close numfile;
=cut
$anum=$ENV{'QUERY_STRING'};

($question, $vara, $varb, $varc, $numa, $numb, $numc) = split(/--/, @vote[$anum]);
if($avar eq "vara"){$numa++};
if($avar eq "varb"){$numb++};
if($avar eq "varc"){$numc++};

open rec, ">text/votes.txt";
$r=0;
while ($r lt $i)
{
  $r++;
  chomp;
  if($r == $anum){
    if($avar eq "varc"){$nulchar="\n"} else {$nulchar=""};
    print rec $question, "--",$vara, "--", $varb, "--", $varc, "--", $numa, "--", $numb, "--", $numc, $nulchar;
  }else{print rec @vote[$r]};
};
close rec;

if($avar eq ""){print "<FONT COLOR=00cccc SIZE=+3>Было бы лучше, если бы Вы указали какой-либо ответ !!!</FONT>
<META HTTP-EQUIV=Refresh CONTENT=0;URL=../../index.htm>";
}else{
print "<FONT COLOR=00FF00 SIZE=+3>Спасибо за указанную информацию !!!</FONT>
<META HTTP-EQUIV=Refresh CONTENT=0;URL=$go_url"};