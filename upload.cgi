#!/usr/bin/perl --

# Use a Perl, PHP or Python uploader to upload your files then feed them into the code below.  Sorry for the kluginess of the code and the lack of comments at the moment...I'm a newbie programmer.
#
# I created a folder with a huge static number folder named "856597900390625" in which to send csv files from each spreadsheet for building into a combined Excel recovery csv.
#
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\856597900390625\\";

opendir (DIR, "$dir/");
@FILES = grep(/.csv/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}
#
# Taking the opportunity to delete previous user files over a day old.  Uploads2 is an old directory I may not use an longer.
#


$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\";

opendir (DIR, "$dir/");
@FILES = grep(/.rtf|.doc|.docx|.docm|.dotx|.dotm|.odt|.ott|.sxw|.stw|.xlsx|.xltx|.xltm|.xlsm|.xlsb|.xlam|.ods|.ots|.sxc|.stc|.pptx |.pptm|.ppsm|.ppsx|.odp|.otp|.sxi|.sti|.txt|.csv/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 1) {
      unlink("$dir/$FILES");
   }
}
# 
# cgi-bin2 is where files are currently stored once uploaded.
#
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\";

opendir (DIR, "$dir/");
@FILES = grep(/.rtf|.doc|.docx|.docm|.dotx|.dotm|.odt|.ott|.sxw|.stw|.xlsx|.xltx|.xltm|.xlsm|.xlsb|.xlam|.ods|.ots|.sxc|.stc|.pptx |.pptm|.ppsm|.ppsx|.odp|.otp|.sxi|.sti|.txt|.csv/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 1) {
      unlink("$dir/$FILES");
   }
}
#
# Beginning of the programming section.  Which starts with a random number generation and creating of error files.
#
		my $random_number = rand();
		my $random_long = $random_number*1000000000000000;
		my $lcsuccess = lc($success[0]);
		my $lcsuccessrand = $random_long.'.'.$lcsuccess;
		rename ($lcsuccess, $lcsuccessrand);
		
		my $errorpath1 = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\errorfiles\\error1.txt";
		my $errorpath2 = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\errorfiles\\error2.txt";		
		my $errorpath3 = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\errorfiles\\error3.txt";		
		my $errorpath4 = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\errorfiles\\error4.txt";		
		my $errorpath5 = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\errorfiles\\error5.txt";
#
# get the extension
# we can assume everything after the last . found is the extension
#
	# my @file_type   = split(/\./, $lcsuccessrand);
	# my $file_type   = $file_type[$#file_type];
#
# libxml2-2.dll	kept disappearing so thus was my way to fix it
#
	use File::Copy;
	my $new_copy = 'libxml2-2.dll';
	my $original = 'social.dll';
	if (-e "libxml2-2.dll") {
		print ("");
		} else {
		copy( $original, $new_copy ) or die "Copy failed: $!";
		}
#
# I use random numbers in the folders and text file names for security reasons: so hackers can't guess the names of files
#

#  
# First the sections here check if the file is a docx one and do appropriate extractions if it is.  
#

print qq~
<table cellspacing="12" cellpadding="12" style="border: 1px solid #000080; border-collapse: collapse;">
<tr>
<td style="vertical-align:top; width: 322px;">
<ul>
	<li><a href="../index.html">Home</a></li>
	<li><a href="../instructions.html">Instructions</a></li>
	</ul>
<center><em>Click <a href="http://www.saveofficedata.com/">here</a> to recover another document.</em></center>
<br/><center>
</td>	
~;
	
#
# The version 1 text file is set up to include a random number
#
	my $saved = $lcsuccessrand.'.alg-version-1.txt';
#
# Text extraction occurs here already using a configuration which fixes the xml. doctotext.exe is available from Silvercoders (http://silvercoders.com/en/products/doctotext) for free.  I use the no-frills unzipper that I commissioned from a programmer with the handle CCY.  He's from China.  Anyway the unzipper is here: http://www.godskingsandheroes.info/software/#No-Frills_Unzipper
#	
 	open my $wfh, "| doctotext.exe --fix-xml --unzip-cmd=\"no-frills.exe %a %d %f\" $lcsuccessrand > \"$saved\" 2> error.txt";
	close $wfh;
	
# This section displays the link to the text file extracted above.
#	
	my $filesize = -s "$saved";
	if ($filesize == 0){
	print qq~<td style = "width: 500px; border: 1px solid #000080; border-collapse: collapse;"><center><br/>
	<br/><b>$lcsuccess</b> has uploaded successfully.</center>
	<br/>The results are:
	<br/><ul><li>The first version algorithm returned no results.</li>~;
	}else{
	print qq~<td style="width: 700px; border: 1px solid #000080; border-collapse: collapse;"><center><br/><b>$lcsuccess</b> has uploaded successfully.</center>
	<br/>The results are:
	<br/><ul><li><a href=\"$saved\" target=\"_blank\"><b>$saved</b></a></li>~;
	}
#
# This section uses the same extractor but the configuration strips the xml tags this time instead of trying to fix them.
#   
	if ($file_type eq "xlsx"){
	print qq~<br/><li>The second version algorithm returned no results</li>~;
	} else {
	my $saved2 = $lcsuccessrand.'.alg-version-2.txt';
	open my $wfh, "| doctotext.exe --strip-xml --unzip-cmd=\"no-frills %a %d %f\" $lcsuccessrand > \"$saved2\" 2> $errorpath2";
	close $wfh;	
		my $filesize2 = -s "$saved2";
	 	if($filesize2 <= 22){
	 	print qq~<br/><li>The second version algorithm returned no results.</li>~;
	  	} else {
	  	print qq~<br/><li><a href=\"$saved2\" target="_blank"><b>$saved2</b></a></li>~;
	}
	}

#`tskill	coffec`;

 	my @file_type   = split(/\./, $lcsuccessrand);
	my $file_type   = $file_type[$#file_type];
#
# coffec.exe is an MS Office docx, xlsx,a nd pptx command line text extractor I commissioned.  It's available as free here: http://www.godskingsandheroes.info/software/#CMD
#
if ($file_type eq "docx"){
	my $saved3 = $lcsuccessrand.'.alg-version-3.txt';
	open my $wfh, "| coffec.exe -t $lcsuccessrand > \"$saved3\" 2> $errorpath3";
	close $wfh;	
	print qq~<br/><li><a href=\"$saved3\" target="_blank"><b>$saved3</b></a></li>~;
} elsif ($file_type eq "pptx"){
	my $saved3 = $lcsuccessrand.'.alg-version-3.txt';
	open my $wfh, "| coffec.exe -t $lcsuccessrand > \"$saved3\" 2> $errorpath3";
	close $wfh;			
	print qq~<br/><li><a href=\"$saved3\" target="_blank"><b>$saved3</b></a></li>~;
} elsif ($file_type eq "xlsx"){
	my $coffecfile = "856597900390625/".$lcsuccessrand;
	copy($lcsuccessrand, $coffecfile);
	my $saved3 = $lcsuccessrand.'.alg-version-3.txt';
	chdir("856597900390625");
	open my $wfh, "| coffec.exe -t $lcsuccessrand > test2.txt 2> $errorpath3";
	close $wfh;	
	@ARGV = <*.csv>;
	open SEL, '>', "$saved3" or die $!;
	while (<>) {
 	print SEL;
	}
	close SEL;
	chdir("../");
	my $saved3coffec = "856597900390625/".$saved3;
	copy($saved3coffec, $saved3);
	print qq~<br/><li><a href=\"$saved3\"target="_blank"><b>$saved3</b></a></li>~;
	


	$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\856597900390625\\xl\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

	$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\856597900390625\\xl\\worksheets\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

	
} else {
	print qq~<br/><li>The third algorithm version returned no results.</li>~;
}

#
# If the file is an Excel one, the next section runs the fix xml configuration of the doctotext extractor against the file.
# Note the file does not run the strip xml configuration as this does not work with xml files.
# Sandeep Kumar wrote a nice docx2txt converter used in this section.  It's avilable here: http://sourceforge.net/projects/docx2txt/
#

if($file_type eq 'docx'){
	my $saved4 = $lcsuccessrand.'.alg-version-4.txt';
	open my $wfh, "| sandeepsconverter.bat $lcsuccessrand > \"$saved4\" 2> $errorpath4";
	close $wfh;	
 	my $lcsuccessrandtxt = $lcsuccessrand.'.txt';
 	rename($lcsuccessrandtxt, $saved4);
	print qq~<br/><li><a href=\"$saved4\" target="_blank"><b>$saved4</b></a></li>~;
#
# I'm using corrupt pptx2txt in this section.  That program is an offshoot of http://sourceforge.net/projects/pptx2txt/ converter by sopan_shewale which in turn is an offshoot of Sandeep Kumar's project.  My project is available here: https://sourceforge.net/p/corruptpptx2txt/home/, and right now I use the no-frills unzipper because  CakeCMD does not end sometimes, backing up the servers I'm on eventually.
#
	
	}elsif ($file_type eq 'pptx'){
	my $saved4 = $lcsuccessrand.'.alg-version-4.txt';
	open my $wfh, "| pptxconverter.exe $lcsuccessrand > \"$saved4\" 2> $errorpath4";
	close $wfh;	
 	my $lcsuccessrandtxt = $lcsuccessrand.'.txt';
 	rename($lcsuccessrandtxt, $saved4);
	print qq~<br/><li><a href=\"$saved4\" target="_blank"><b>$saved4 (Includes possible Notes text recovery.)</b></a></li>~;
}else {
	print qq~<br/><li>The fourth algorithm version returned no results.</li>~;
}
#
# This section is a new feature that removes part of a key Open Office XML file which once done will often allow the opening of corrupt files that refused to open before.
#
if($file_type eq 'odt'){

	`no-frills $lcsuccessrand openofficewriter`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficewriter') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.odt';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5odt = $lcsuccessrand.'.alg-version-5.odt';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5odt`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} elsif ($file_type eq 'ott'){

	`no-frills $lcsuccessrand openofficewriter`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficewriter') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.ott';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5ott = $lcsuccessrand.'.alg-version-5.ott';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5ott`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} elsif ($file_type eq 'ods'){

	`no-frills $lcsuccessrand openofficecalc`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficecalc') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.ods';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5ods = $lcsuccessrand.'.alg-version-5.ods';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5ods`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} elsif ($file_type eq 'ots'){

	`no-frills $lcsuccessrand openofficecalc`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficecalc') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.ots';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5ots = $lcsuccessrand.'.alg-version-5.ots';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5ots`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} elsif ($file_type eq 'odp'){

	`no-frills $lcsuccessrand openofficeimpress`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficeimpress') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.odp';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5odp = $lcsuccessrand.'.alg-version-5.odp';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5odp`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} elsif ($file_type eq 'otp'){

	`no-frills $lcsuccessrand openofficeimpress`;
	
	use File::Copy;
	$filetobecopied = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\executables2\\manifest.xml";
	$newfile = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\META-INF\\manifest.xml";
	copy($filetobecopied, $newfile) or die "File cannot be copied.";
	
	chdir('openofficeimpress') or die "$!";
	my $zipper = "c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\7za.exe";
	my $saved5 = "..\\".$lcsuccessrand.'.alg-version-5.otp';
	`$zipper a -tzip $saved5 \*`;
	
	chdir("c:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\") or die "$!";
	my $saved5otp = $lcsuccessrand.'.alg-version-5.otp';
	my $saved5zip = $lcsuccessrand.'.alg-version-5.zip';
	`7za.exe a -tzip $saved5zip $saved5otp`;
	print qq~<br/><li><a href=\"$saved5zip\" target="_blank"><b>$saved5zip</b></a></li>~;
} 
print qq~</ul><center>
</center>
</td>
</tr>
	</table>
~;
	
# The section below removes user files for a little better security, though I guess it is rather primitive.
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\word\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\word\\_rels\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}


$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\ppt\\slides\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\xl\\worksheets\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\xl\\drawings\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

opendir (DIR, "$dir/");
@FILES = grep(/.vml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\xl\\drawings\\_rels\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\xl\\media\\";


opendir (DIR, "$dir/");
@FILES = grep(/image/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\word\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\word\\_rels\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}


$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\ppt\\slides\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\xl\\worksheets\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\xl\\drawings\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

opendir (DIR, "$dir/");
@FILES = grep(/.vml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\xl\\drawings\\_rels\\";


opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\uploads2\\xl\\media\\";


opendir (DIR, "$dir/");
@FILES = grep(/image/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\META-INF\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\Thumbnails\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficecalc\\Pictures\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}


$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\META-INF\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\Thumbnails\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficeimpress\\Pictures\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\Pictures\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\Thumbnails\\";

opendir (DIR, "$dir/");
@FILES = grep(/.png/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\openofficewriter\\META-INF\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xml/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\";

opendir (DIR, "$dir/");
@FILES = grep(/CGItemp*/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

$dir = "cgi-bin2\\856597900390625\\";

opendir (DIR, "$dir/");
@FILES = grep(/.csv/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\856597900390625\\";

opendir (DIR, "$dir/");
@FILES = grep(/.txt/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}
$dir = "C:\\inetpub\\vhosts\\saveofficedata.com\\httpdocs\\cgi-bin2\\856597900390625\\";

opendir (DIR, "$dir/");
@FILES = grep(/.xlsx/,readdir(DIR));
closedir (DIR);

## DELETE THE .TXT FILES THAT ARE OLDER THAN 1 DAY

foreach $FILES (@FILES) {
   if (-M "$dir/$FILES" > 0) {
      unlink("$dir/$FILES");
   }
}

