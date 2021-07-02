sub datetosec {
use Time::Local;
my $gdate = $_[0];
($day,$month,$year) = ($gdate =~ /(\d\d?):(\d\d?):(\d\d\d\d)/);
my $gtime = timelocal($sec,$min,$hour,$day,$month-1,$year-1900);
return ($gtime);
}

sub fastvalues {
my $fname = $_[0];
my @res;
my @val;
open (FILE, "<$fname") or &fastoutprint ($cant_open.'&nbsp;'.$fname);
while (<FILE>) {
@val = split ("~", $_);
push(@res, @val); 
}
close (FILE);
return (@res);
}

sub fastread {
my $fname = $_[0];
open (FILE, "<$fname") or &fastoutprint ($cant_open.'&nbsp;'.$fname); 
undef $/; 
my $res = <FILE>; 
close(FILE);
return ($res); 
}

sub fastadd {
my $fname = $_[0];
my $add = $_[1];
open (FILE, ">>$fname") or &fastoutprint ($cant_open.'&nbsp;'.$fname); 
print FILE "$add~";
close(FILE);
}

sub ssi {
$tests = $_[0];
while ($tests =~ /(.!--.include virtual.+?-->)/gi) {
$changes = $1;
$filenames=($changes =~ /\"(.+?)\"/)[0];
$firsts = substr($filenames,0,1);
if ($firsts eq '/') {
$includes = fastread ($basedir.$filenames); 
} else {
$includes = fastread ($fhdir.$filenames);
}
$tests =~ s/$changes/$includes/;
}
return $tests;
}

sub golosov {
my $gol = $_[0];
my $res;
if ($gol =~ /1\d$/) {
$res = $right_vote3;
} elsif ($gol =~ /[234]$/) {
$res = $right_vote2;
} elsif ($gol =~ /1$/) {
$res = $right_vote1;     
} else {
$res = $right_vote3;
}
return ($res);
}

sub fastoutprint {
my $gprint = $_[0];
print "Content-type: text/html\n\n";
print $head;
print $gprint;
print qq~<p align="center"><font size="-1">Scripted by <a href="http://www.script.ee" target="_new">Scripting Project</a></font></p>~;
print $foot;
exit;
}

sub adminprint {
my $gprint = $_[0];
print "Content-type: text/html\n\n";
print $head;
if (!$menuneed) {
$supermenu .= qq~<a href="$admin_file_url?main">$home</a><br><br>~;
print $supermenu;
}
print $gprint;
print qq~<p align="center"><font size="-1">Scripted by <a href="http://www.script.ee" target="_new">Scripting Project</a></font></p>~;
print $foot;
exit;
}

1;
