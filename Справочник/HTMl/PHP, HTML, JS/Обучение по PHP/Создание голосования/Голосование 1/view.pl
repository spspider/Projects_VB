#!/usr/bin/perl
print "Content-Type: text/html\n\n";

open votes, "<text/votes.txt";
$i=0;
while (<votes>)
{
$i++;
chomp;
@vote[$i] = $_;
($question, $vara, $varb, $varc, $numa, $numb, $numc) = split(/--/, @vote[$i]);
print "$question<br>- $vara (".$numa.")<br> - $varb (".$numb.")<br> - $varc (".$numc.")<hr>";
}
close votes;