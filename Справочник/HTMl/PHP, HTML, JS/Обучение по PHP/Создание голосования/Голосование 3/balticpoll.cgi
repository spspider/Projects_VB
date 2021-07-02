#!/

#Скрипт голосования был разработан
#by Haze (haze@script.ee) для Scripting Project (www.script.ee)
#Изменение и последующая продажа скрипта запрещена без разрешения автора!
#Данная версия является freeware

require "config.pl";
require "functions.pl";
require "words.pl";

$head = fastread ($fhdir.$header);
$head = ssi($head);
$foot = fastread ($fhdir.$footer);
$foot = ssi ($foot);

if ((!$ENV{'CONTENT_LENGTH'}) && (!$ENV{'QUERY_STRING'})) {
fastoutprint ($error_request);
}


$out_results = $ENV{'QUERY_STRING'};
if ($out_results =~ /\d/) {
$input{counter} = $out_results;
$out_results = 'yes';
}

if ($out_results ne 'yes') {
read(STDIN, $buffer, $ENV{'CONTENT_LENGTH'});
@pairs = split(/&/, $buffer);
foreach $pair (@pairs) {
  ($name, $value) = split(/=/, $pair);
  $value =~ tr/+/ /;
  $value =~ s/%([a-fA-F0-9][a-fA-F0-9])/pack("C", hex($1))/eg;
  $value =~ s/<!--(.|\n)*-->//g;
  $input{$name} = $value;
}

if ($ENV{'HTTP_X_FORWARDED_FOR'}){
$ip=$ENV{'HTTP_X_FORWARDED_FOR'};
} else { $ip=$ENV{'REMOTE_ADDR'};
}
}

$dirname = $scriptdir.$input{counter}.'/';
$time = time();
$time = $time + ($timeshift*3600);

$startsec = datetosec ($starttime);
if ($endtime eq "endless") {
$endsec = $time;
} else {
$endsec = datetosec ($endtime);
}

opendir (DIR, $dirname) or &fastoutprint ($cant_open.'&nbsp;'.$dirname);
while (defined($file = readdir(DIR))) {
next if $file !~ /\.txt/;
push (@pollfiles, $file);
}

if ($out_results ne 'yes') {
@ip_base = fastvalues ($dirname."0.txt");
if ($ip_test eq 'Yes') {
foreach $ip_from_base (@ip_base) {
if ($ip eq $ip_from_base) {
fastoutprint ($ip_alert);
exit;
}
}
}
}

if ($out_results ne 'yes') {
foreach $ip_from_base (@ip_base) {
if ($ip eq $ip_from_base) {
$testip = 'Yes';
}
}
if ($testip ne 'Yes') {
fastadd ($dirname."0.txt", $ip);
}
}

@data = fastvalues ($dirname."data.dat");
for ($i=1; $i<=scalar(@data); $i++) {
$max = length ($data[$i]);
if ($max > $width2) {
$width2 = $max;
}
}
$width2 = $width2*7 + 90;

if ($out_results ne 'yes') {
if ($input{vote}) {
fastadd ($dirname.$input{vote}.".txt", $time);
} else {
for ($i = 1; $i <= ($#pollfiles + 1); $i++) {
if ($input{$i}){
fastadd ($dirname.$i.".txt", $time);
}
}
}
}

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


$outprint = $begintmp;

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
fastoutprint ($outprint);