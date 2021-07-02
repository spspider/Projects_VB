#!/

#Скрипт голосования был разработан
#by Haze (haze@script.ee) для Scripting Project (www.script.ee)
#Изменение и последующая продажа скрипта запрещена без разрешения автора!
#Данная версия является freeware

require "config.pl";
require "functions.pl";
require "words.pl";

$admin_file_url = $scripturl.'admin.cgi';
$admin_file = $scriptdir.'admin.cgi';

$cdata = $ENV{'HTTP_COOKIE'};
$autopass = ($cdata =~ /BALPASS=(\w+)/)[0];

if ($ENV{'CONTENT_LENGTH'}) {
read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});
@pairs = split(/&/, $buffer);
foreach $pair (@pairs) {
  ($name, $value) = split(/=/, $pair);
  $value =~ tr/+/ /;
  $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;
  $input{$name} = $value;
}
}

if ($admin_file ne $ENV{SCRIPT_FILENAME}){
$admin_file_url = $ENV{SCRIPT_NAME};
if ($input{setuppoll}) {
$input{pass} = $password;
$install = 1;
&setup_poll ();
} else {
$menuneed = 'No';
&setup_poll_print ();
}
}

$head = fastread ($fhdir.$header);
$head = ssi($head);
$foot = fastread ($fhdir.$footer);
$foot = ssi ($foot);

opendir (DIR, $scriptdir) or &adminprint ($cant_open.'&nbsp;'.$scriptdir);
while (defined($file = readdir(DIR))) {
next if $file !~ /^(\d+)$/;
push (@allpolls, $file);
$last_poll = $file if $file > $last_poll;
}

if (($ENV{'QUERY_STRING'}) || ($ENV{'CONTENT_LENGTH'})) {
if ($ENV{'QUERY_STRING'} eq 'add') {
&add_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'delete') {
&delete_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'edit') {
&edit_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'code') {
&code_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'config') {
&setup_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'main') {
&enter_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'annul') {
&annul_poll_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'tmp') {
&edit_templates_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'words') {
&words_print ();
} elsif ($ENV{'QUERY_STRING'} eq 'results') {
&results_print ();
}
}

if ($input{addpoll}) {
&add_poll ();
} elsif ($input{deletepoll}) {
&delete_poll ();
} elsif ($input{code}) {
&code_poll ();
} elsif ($input{installpoll}) {
&setup_poll_print ();
} elsif ($input{enter}) {
&enter_poll_print ();
} elsif ($input{setuppoll}) {
&setup_poll ();
} elsif ($input{annulpoll}) {
&annul_poll ();
} elsif ($input{editpoll}) {
&edit_poll ();
} elsif ($input{editpollsuc}) {
&edit_pollsuc ();
} elsif ($input{edittmp}) {
&edit_templates_print ();
} elsif ($input{edittmpsuc}) {
&edit_templates ();
} elsif ($input{wordschange}) {
&words ();
} elsif ($input{resultspoll}) {
&results_poll ();
} elsif ($input{results}) {
&results_poll ();
} elsif ($input{changepoll}) {
&change_file_poll ();
}

&new_enter();

sub new_enter {
$menuneed = 'No';
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="enter" value="1">
<table><tr><td>$check_pass:&nbsp;</td><td>
<input TYPE="PASSWORD" class=forma name="pass"></td></tr><tr><td align="center" colspan="2"><br>
<input type="submit" class=forma name="post" value="$enterbut"></td></tr></table></form>~;
adminprint ($outprint);
}

sub words {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&words_print ();
}

$text = "";
$i = 0;
open (FILE, "<words.pl") or fastoutprint ($cant_open.'&nbsp;words.pl');
$/ = "\n";
while (<FILE>) {
$begin = ($_ =~ /(\$\w+?.+?qq\~)/)[0];
$text .= "$begin$input{$i}~;\n";
$i++;
}
close(FILE);

open (FILE, ">words.pl");
print FILE $text;
close (FILE);
adminprint ($words_change_success);
}

sub words_print {
$text = fastread ($scriptdir.'words.pl');
$outprint .= qq~<form action="$admin_file_url" method="POST"><INPUT TYPE="hidden" name="wordschange" value="1"><table>~;
$i = 0;
while ($text =~ /qq\~(.+?)\~/g) {
$word = $1;
$outprint .= qq~<tr><td><textarea class=forma name="$i" rows=3 cols=60>$word</textarea><br></td></tr>~;
$i++;
}
$outprint .= qq~<tr><td><br>$pass:&nbsp;<input TYPE="PASSWORD" class=forma name="pass" value="$autopass">
</td></tr><tr><td align="center"><br><input type="submit" class=forma name="post" value="$changebut"></td></tr></table></form>~;
adminprint ($outprint);
}

sub edit_templates {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&edit_templates_print ();
}

if ($input{tmpname}) {
$file = $scriptdir.$input{tmpname};
} else {
$ouprint = $file_forgot;
&edit_templates_print ();
}

open (FILE, ">$file");
print FILE $input{intmp};
close (FILE);
adminprint ($tmp_success);
exit;
}

sub edit_templates_print {
if ($input{whichtmp}) {
$datatmp = fastread ($fhdir.$input{whichtmp});
} elsif ($input{intmp}) {
$datatmp = $input{intmp};
}


$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="edittmp" value="1">
<table><tr><td>$hinttmp</td></tr>
<tr><td><br><select name="whichtmp">
<option value="$header">$headtmp
<option value="$footer">$foottmp
<option value="$resulttmp">$restmp
</select>&nbsp;&nbsp;<input type="submit" class=forma name="post" value="$openbut">
</td></tr></table></form>~;

$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="edittmpsuc" value="1">
<INPUT TYPE="hidden" name="tmpname" value="$input{whichtmp}">
<table><tr><td><br><textarea class=forma name="intmp" rows=30 cols=60>$datatmp</textarea>
</td></tr><tr><td><br>$pass:&nbsp;<input TYPE="PASSWORD" class=forma name="pass" value="$autopass">
</td></tr><tr><td align="center"><input type="submit" class=forma name="post" value="$savebut">
</td></tr></table></form>~;

adminprint ($outprint);
}


sub enter_poll_print {
$menuneed = 'No';
if ($input{enter}) {
if ($input{pass} ne $password) {
$outprint = qq~$pass_error~;
&new_enter();
}
}

@amenu = ('add','delete','edit','annul','code','results','tmp','words','config');
if ($input{enter}) {
print "Set-Cookie: BALPASS=$password\n";
}
$outprint .= qq~<UL>~;
for ($i = 0; $i <= 8; $i++) {
$file = $admin_file_url.'?'.$amenu[$i];
if ($last_poll != 0) {
$outprint .= qq~<LI><a href="$file">$menu[$i]</a>~;
} else {
if (($i < 1) || ($i > 5)) {
$outprint .= qq~<LI><a href="$file">$menu[$i]</a>~;
} else {
$outprint .= qq~<LI>$menu[$i]~;
}
}
}

$outprint .= qq~</UL>~;


adminprint ($outprint);
}

sub edit_pollsuc {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&edit_poll ();
}
$file = $scriptdir.$input{wheditpoll}.'/data.dat';
for ($i=0; $i <= $input{wheditpollv}; $i++) {
$text = $text.$input{$i}.'~';
}
open (FILE, ">$file");
print FILE $text;
close (FILE);
adminprint ($edit_success);
}

sub edit_poll {
$supermenu = qq~<a href="$admin_file_url?edit">$menu[2]</a> >> ~;
if ($input{pass} ne $password) {
$outprint = $pass_error;
&edit_poll_print ();
}
$file = $scriptdir.$input{edit}.'/data.dat';
@values = fastvalues($file);
$scal = $#values;
if (!$input{0}) {
for ($i=0; $i < scalar(@values); $i++) {
$input{$i} = $values[$i];
}
}

$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="editpollsuc" value="1">
<INPUT TYPE="hidden" name="wheditpoll" value="$input{edit}">
<INPUT TYPE="hidden" name="wheditpollv" value="$scal">
<table>
<tr><td>$question:&nbsp;</td><td><input name="0" class=forma size="50" value="$input{0}"></td></tr>
~;

for ($i=1; $i < scalar(@values); $i++) {
$outprint .= qq~<tr><td>$answer&nbsp;$i:&nbsp;</td><td><input name="$i" class=forma size="30" value="$input{$i}"></td></tr>~;
}

$outprint .= qq~<tr><td><br>$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass">
</td></tr><tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$editbut"></table></form>~;

adminprint ($outprint);
}

sub annul_poll {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&annul_poll_print ();
}
$dir = $scriptdir.$input{annul}.'/';
opendir (DIR, $dir) or &adminprint ($cant_open.'&nbsp;'.$dir);
while (defined($file = readdir(DIR))) {
if ($file =~ /.txt/) {
$nfile = $scriptdir.$input{annul}.'/'.$file;
open (FILE, ">$nfile") or &adminprint ($cant_open.'&nbsp;'.$fname);
print FILE '~';
close (FILE);
}
}
adminprint ($annul_success);
}

sub annul_poll_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="annulpoll" value="1">
<table><tr><td>$annul:&nbsp;</td><td><select name="annul">~;

foreach $poll (@allpolls) {
my $file = $scriptdir.$poll.'/data.dat';
my @data = fastvalues ($file);
$outprint .= qq~<option value="$poll">$data[0]~;
}
$outprint .= qq~</select></td></tr><tr><td><br>
$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$annulbut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub edit_poll_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="editpoll" value="1">
<table><tr><td>$edit:&nbsp;</td><td><select name="edit">~;

foreach $poll (@allpolls) {
my $file = $scriptdir.$poll.'/data.dat';
my @data = fastvalues ($file);
$outprint .= qq~<option value="$poll">$data[0]~;
}
$outprint .= qq~</select></td></tr><tr><td><br>
$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$editbut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub setup_poll {
if ($input{newpass}) {
if ($input{newpass} ne $input{newpasscon}) {
$outprint = qq~$pass_confirm_error~;
&setup_poll_print ();
} else {
$input{$#descript+1} = $input{newpass};
print "Set-Cookie: BALPASS=$input{newpass}\n";
}
} else {
$input{$#descript+1} = $password;
}

if ($install) {
$input{$#descript+1} = $password;
}

if (($input{pass} ne $password) && ($install ne '1')) {
$outprint = qq~$pass_error~;
&setup_poll_print ();
}

if (($input{12} !~ /^(\d\d?):(\d\d?):(\d\d\d\d)$/) || (($input{13} !~ /^(\d\d?):(\d\d?):(\d\d\d\d)$/) && ($input{13} ne 'endless'))) {
$outprint = qq~$date_error~;
&setup_poll_print ();
}

$text = "";
$i = 0;
open (FILE, "<config.pl") or fastoutprint ($cant_open.'&nbsp;config.pl');
$/ = "\n";
while (<FILE>) {
$begin = ($_ =~ /(\$\w+?.+?qq\~)/)[0];
$text .= "$begin$input{$i}~;\n";
$i++;
}
close(FILE);

open (FILE, ">config.pl") or &adminprint ($cant_open.'&nbsp;config.pl');
print FILE $text;
close (FILE);
adminprint ($config_change_success);
}

sub setup_poll_print {
if (!$input{setuppoll}) {
$text = fastread ('config.pl');
$i = 0;
while ($text =~ /qq\~(.*?)\~/g) {
$input{$i} = $1;
$i++;
}
}


$dir = $ENV{'SCRIPT_FILENAME'};
@direc = split ('/', $dir);
$dir = join ("/", @direc[0 .. $#direc-1]);
$dir = $dir.'/';

$outprint .= qq~<form action="$admin_file_url" method="POST"><INPUT TYPE="hidden" name="setuppoll" value="1">
<table><tr><td>$options</td></tr>
<tr><td>$script_dir&nbsp;$dir</font></td></tr>
<tr><td><input name="0" class=forma size="50" value="$input{0}"></td></tr>~;

for ($i=1; $i <= $#descript; $i++) {
$outprint .= qq~<tr><td><br>$descript[$i]</td></tr>
<tr><td><input name="$i" class=forma size="50" value="$input{$i}"></td></tr>~;
}

$outprint .= qq~<tr><td><br>$newpass</td></tr>
<tr><td><input class=forma TYPE="PASSWORD" name="newpass"></td></tr>
<tr><td><br>$newpasscon</td></tr>
<tr><td><input class=forma TYPE="PASSWORD" name="newpasscon"></td></tr>
<tr><td><br><br>$pass:&nbsp;<input class=forma TYPE="PASSWORD" name="pass" value="$autopass"></td></tr>
<tr><td align="center"><br>
<input type="submit" class=forma name="post" value="$changebut">
</td></tr></table></form>
~;
adminprint ($outprint);
}

sub add_poll {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&add_poll_print ();
}
if (!$input{0}) {
$outprint = $add_error;
&add_poll_print ();
}
$newdirname = ($last_poll+1);
mkdir $newdirname, 0755;

$newdirname = $scriptdir.$newdirname.'/';

$data_add = $input{0}.'~';
$file = $newdirname.'0.txt';
open (FILE, ">$file") or &adminprint ($cant_open.'&nbsp;'.$file);
close (FILE);

for ($i = 1; $i <= $maxvariant; $i++) {
if (!$input{$i}) {
last;
} else {
$data_add .= $input{$i}.'~';
$file = $newdirname.$i.'.txt';
open (FILE, ">$file") or &adminprint ($cant_open.'&nbsp;'.$file);
close (FILE);
}
}
$file = $newdirname.'data.dat';
open (FILE, ">$file") or &adminprint ($cant_open.'&nbsp;'.$file);
print FILE $data_add;
close (FILE);
adminprint ($add_poll_success);
}

sub delete_poll {
if ($input{pass} ne $password) {
$outprint = $pass_error;
&delete_poll_print ();
}
$dirname = $scriptdir.$input{'delete'}.'/';
opendir (DIR, $dirname) or &adminprint ($cant_open.'&nbsp;'.$dirname);
while (defined($file = readdir(DIR))) {
if (($file =~ /\.txt/) || ($file =~ /\.dat/)) {
push (@pollfiles, $dirname.$file);
}
}
unlink(@pollfiles) == @pollfiles or &adminprint ($delete_error);
rmdir $input{'delete'} or &adminprint ($delete_error);
adminprint ($delete_poll_success);
}

sub add_poll_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="addpoll" value="1">
<table border="0">
<tr><td>$question:&nbsp;</td><td><input name="0" class=forma size="50" value="$input{0}"></td></tr>
~;

for ($i=1; $i <= $maxvariant; $i++) {
$outprint .= qq~<tr><td>$answer&nbsp;$i:&nbsp;</td><td><input name="$i" class=forma size="30" value="$input{$i}"></td></tr>~;
}

$outprint .= qq~<tr><td><br>$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass">
</td></tr><tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$addbut"></table></form>~;

adminprint ($outprint);
}

sub code_poll_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="getcode" value="1">
<table><tr><td>$get_code:&nbsp;</td><td><select name="code">~;

foreach $poll (@allpolls) {
my $file = $scriptdir.$poll.'/data.dat';
my @data = fastvalues ($file);
$outprint .= qq~<option value="$poll">$data[0]~;
}
$outprint .= qq~</select></td></tr><tr><td colspan="2">
<input type="radio" name="radiocheck" value="radio" checked>$radio<br>
<input type="radio" name="radiocheck" value="check">$check
</td></tr><tr><td><br>
$pass:&nbsp;</td><td><br><input class=forma TYPE="PASSWORD" name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$getbut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub code_poll {
$supermenu = qq~<a href="$admin_file_url?code">$menu[4]</a> >> ~;
if ($input{pass} ne $password) {
$outprint = $pass_error;
&code_poll_print ();
}
$outprint = qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="changepoll" value="1"><table><tr><td>$insert_this_code<td></tr>~;
my $file = $scriptdir.$input{code}.'/data.dat';
my @data = fastvalues ($file);
my $action = $scripturl.'balticpoll.cgi';
$outprint .= qq~<tr><td><br>
<textarea class=forma cols="80" rows="15" name="pollcodetext"><form action="$action" method="POST">$data[0]<br><br>\n~;
$outprint .= qq~<INPUT TYPE="hidden" name="counter" value="$input{code}">\n~;
if ($input{radiocheck} eq 'radio') {
$outprint .= qq~<input type="radio" name="vote" value="1" checked>$data[1]<br>\n~;
for ($i = 2; $i <= $#data; $i++) {
$outprint .= qq~<input type="radio" name="vote" value="$i">$data[$i]<br>\n~;
}
} else {
for ($i = 1; $i <= $#data; $i++) {
$outprint .= qq~<input TYPE="CHECKBOX" name="$i" value="yes">$data[$i]<br>\n~;
}
}
$outprint .= qq~<br><input type="submit" class=formad value="$vote"></form>\n<a class=menua href="$action?$input{code}">$results</a></textarea>~;
$outprint .= qq~</td></tr></table><table><tr><td colspan="2">$change_poll_in_file</td></tr><tr><td><br>
$file_with_poll:&nbsp;</td><td><br><input name="outputfile" class=forma size="50" value="$outputfile"></td></tr><tr><td>
$pass:&nbsp;</td><td><input class=forma TYPE="PASSWORD" name="pass" value="$autopass"></td></tr>
<tr><td align="center" colspan="2"><br><input type="submit" class=forma name="post" value="$changebut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub delete_poll_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="deletepoll" value="1">
<table><tr><td>$delete_poll:&nbsp;</td><td><select name="delete">~;

foreach $poll (@allpolls) {
my $file = $scriptdir.$poll.'/data.dat';
my @data = fastvalues ($file);
$outprint .= qq~<option value="$poll">$data[0]~;
}
$outprint .= qq~</select></td></tr><tr><td><br>
$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$deletebut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub results_print {
$outprint .= qq~<form action="$admin_file_url" method="POST">
<INPUT TYPE="hidden" name="resultspoll" value="1">
<table><tr><td>$results_poll:&nbsp;</td><td><select name="results">~;

foreach $poll (@allpolls) {
my $file = $scriptdir.$poll.'/data.dat';
my @data = fastvalues ($file);
$outprint .= qq~<option value="$poll">$data[0]~;
}
$outprint .= qq~</select></td></tr><tr><td><br>
$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br>
<input type="submit" class=forma name="post" value="$nextbut">
</td></tr></table></form>~;
adminprint ($outprint);
}

sub results_poll {
$supermenu = qq~<a href="$admin_file_url?results">$menu[5]</a> >> ~;
if ($input{pass} ne $password) {
$outprint = $pass_error;
if ($input{results}) {
$startsec = datetosec('16:11:2000');
$endsec = time();
$input{startsec} = '16:11:2000';
($day, $month, $year) = (localtime)[3,4,5];
$month = $month + 1;
$year = $year + 1900;
$input{endsec} = "$day:$month:$year";
} else {
&results_print;
}
}

if ((!$input{startsec}) || (!$input{endsec})) {
$startsec = datetosec('16:11:2000');
$endsec = time();
$input{startsec} = '16:11:2000';
($day, $month, $year) = (localtime)[3,4,5];
$month += 1;
$year += 1900;
$input{endsec} = "$day:$month:$year";
} else {
if (($input{startsec} =~ /^(\d\d?):(\d\d?):(\d\d\d\d)$/) || ($input{endsec} =~ /^(\d\d?):(\d\d?):(\d\d\d\d)$/)) {
$startsec = datetosec($input{startsec});
$endsec = datetosec($input{endsec});
} else {
$startsec = datetosec('16:11:2000');
$endsec = time();
$outprint = $date_error;
}
}

$outprint .= qq~
<form action="$admin_file_url" method="POST"><INPUT TYPE="hidden" name="results" value="1">
<table><tr><td>$results_description<br></td></tr></table>
<table><tr><td>
$startsec_description:&nbsp;</td><td><input name="startsec" class=forma value="$input{startsec}">
</td></tr><tr><td>
$endsec_description:&nbsp;</td><td><input name="endsec" class=forma value="$input{endsec}">
<tr><td><br>$pass:&nbsp;</td><td><br><input TYPE="PASSWORD" class=forma name="pass" value="$autopass"></td></tr>
<tr><td colspan="2" align="center"><br><input type="submit" class=forma name="post" value="$nextbut"></td></tr>
</td></tr></table></form>~;

$dirname = $scriptdir.$input{results}.'/';
opendir (DIR, $dirname) or &fastoutprint ($cant_open.'&nbsp;'.$dirname);
while (defined($file = readdir(DIR))) {
next if $file !~ /\.txt/;
push (@pollfiles, $file);
}

@data = fastvalues ($dirname."data.dat");
for ($i=1; $i<=scalar(@data); $i++) {
$max = length ($data[$i]);
if ($max > $width2) {
$width2 = $max;
}
}
$width2 = $width2*7 + 90;

for ($i = 1; $i <= ($#pollfiles); $i++) {
my @value = fastvalues ($dirname.$i.".txt");
for ($t = 0; $t < (scalar (@value)); $t++) {
if (($value[$t] >= $startsec) && ($value[$t] <= ($endsec+86403))) {
push (@this_votes, $value[$t]);
}
}
if (scalar(@this_votes) != 0) {
$votes[0] = $votes[0] + scalar(@this_votes);
$votes[$i] = scalar(@this_votes);
} else {
$votes[0] = $votes[0];
$votes[$i] = 0;
}
@this_votes = ();
}

if ($multicolor =~ /\./) {
foreach (@pollfiles) {
push (@images, $imageurl.$multicolor);
}
} else {
opendir (IMAGEDIR, $imagedir) or &fastoutprint ($cant_open.'&nbsp;'.$dirname);
while (defined($file = readdir(IMAGEDIR))) {
if (($file =~ /\.gif/) || ($file =~ /\.png/) || ($file =~ /\.jpg/)) {
push (@images, $imageurl.$file);
}
}
}
while (scalar(@images) < scalar(@pollfiles)) {
push(@images, @images);
}

$votetmp = fastread ($fhdir.$resulttmp);
$votetmp = ssi($votetmp);
$begintmp = ($votetmp =~ /^(.+)<%start%>/gs)[0];
$medtmp = ($votetmp =~ /<%start%>(.+)<%end%>/gs)[0];
$endtmp = ($votetmp =~ /<%end%>(.+)$/gs)[0];

$poll_header = $data[0];
if ($votes[0] != 0) {
$all_votes_count = $votes[0];
} else {
$all_votes_count = 0;
}
$begintmp =~ s/\$(\w+)/${$1}/g;
$endtmp =~ s/\$(\w+)/${$1}/g;


$outprint .= $begintmp;

for ($v = 1; $v < scalar (@votes); $v++) {
$votes[0] = 1 if $votes[0] == 0;
$width = sprintf ("%.2f", $votes[$v]*100/$votes[0]);
$orfo = golosov($votes[$v]);
$answer = $data[$v];
$amaunt_votes = $votes[$v];
$image = $images[$v-1];
$circle = $medtmp;
$circle =~ s/\$(\w+)/${$1}/g;
$outprint .= $circle;
}

$outprint .= $endtmp;
adminprint ($outprint);
}

sub change_file_poll {
if ($input{pass} ne $password) {
$supermenu = qq~<a href="$admin_file_url?code">$menu[4]</a> >> ~;
adminprint ($pass_error);
}

$text = fastread ($input{outputfile});
if ($text =~ /<!-- Baltic Poll Start -->/) {
$text =~ s/<!-- Baltic Poll Start -->.*<!-- Baltic Poll End -->/<!-- Baltic Poll Start -->$input{pollcodetext}<!-- Baltic Poll End -->/is;
open (FILE, ">$input{outputfile}") or &adminprint ($cant_open.'&nbsp;'.$input{outputfile});
print FILE $text;
close (FILE);
adminprint ($poll_change_success);
} else {
adminprint ($poll_change_error);
}
}