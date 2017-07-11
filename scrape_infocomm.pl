#!/usr/local/bin/perl

# Get a list of vendors presenting at the infocomm tradeshow

use strict;
use DBI;
use LWP::Simple;
use Spreadsheet::WriteExcel;
use Mail::Sender::Easy qw(email);


my (
    $html,
	$html1,
	$html2,
	$html3,
	$url,
	$categories,
	$hrefclose,
	$line,
	$lineiwant,
	$line2,
	$hrefopen,
	@lines,
	@lines2,
	@linessec2,
	$linesec2,
	$linesec2line,
	$hrefopen2,
	$hrefclose2,
	$lineaddwww,
	$lineaddwww2,
	$linesec22,
	$subcategories,
	$filename,
	$workbook,
	$row,
	@rowdata,
	$worksheet,
	$street,
	$addresses,
	@linessec3,
	$number,
	$col,
	$i,
	$name,
	$namejunk,
	$junk,
	$count,
	$cityzipjunk,
	$cityzip,
	$streetjunk,
	$phone,
	$phonejunk,
	$headerformat,
	$pagenoline,
	@pagenolines,
	$pageaddress,
	$pagehrefopen,
	$pagenoline2,
	$crap,
	$crap1,
	$pagenolist,
	$pagecount,
	$centerchunk,
	$stuff_i_want,
	$slurping,
	$pageline,
	$pages,
	@pagelines,
	$topchunk,
	$pagelines,
	$location,
	$locationjunk,
	$phonejunk,
	$phone,
	$name1,
	$city, 
	$state, 
	$zip,
	$statezip,
	$yellowname,
	$yellowstuff_i_want,
	$html5,
	$html6,
	$html7,
	$street,
	$streetjunk,
	$yellowslurping,
	$yellowpagelines,
	$yellowpageline,
	@yellowpagelines,
	$yellowhtml,
	$url2,
	$yellowcity,
	$html8,
	$streetjunk2,
	$actualstreet, 
	$crappola,
	$officialcity,
	$cityline,
	$cityslurp,
	$cityline_i_want
);

$ARGV[0] = $ENV{USERNAME} unless $ARGV[0]; 	

$filename="infocomm.xls";	
	
$workbook = Spreadsheet::WriteExcel->new($filename);
$worksheet = $workbook->add_worksheet();

$row=0;	
$col=0;

$headerformat = $workbook->add_format();
$headerformat->set_bold();
$headerformat->set_underline();

# @rowdata = ("Company", "Location", "Phone Number");
       # $worksheet->write_row($row, 0, \@rowdata, $headerformat);
       # $row++;



#print "start\n";
$pagecount=0;
while ($pagecount < 574)
{	

	$url="http://www.infocommshow.org/networknow/public/nz_ALExhibitorList.aspx";
        $html = get("$url");     
	
	$slurping = 0;
	@pagelines = split /\n/, $html;   #separate each line of the page into @pagelines
	foreach $pageline (@pagelines)   # for every line in the page,
    {
        if ($pageline =~ m/class="listingSummary"/)
		{
		    $slurping = 1;
		}
		elsif ($pageline =~ m/<td colspan="2" style="padding: 0 0 0 15px">/)
		{
		     $slurping = 0;
			 $html=$stuff_i_want;
			 ($html1, $namejunk) = (split /<h1 style="margin:9px 5px 0">/, $html);
			 ($name, $html3) = (split /<span style="border:0; float: right; width: 200px; text-align: right;">/, $namejunk);
			 $name =~ s/&amp;/&/;
			 $name =~ s/^\s*(.*)?/$1/;  # kill leading spaces
                         $name =~ s/\s+$//;  # kill trailing spaces
			 			  
			 $html=$stuff_i_want;
			 ($html1, $locationjunk) = (split /<h3>/, $html);
			 ($location, $html3) = (split /<\/h3>/, $locationjunk);
			 
			 $html=$stuff_i_want;
			 ($html1, $phonejunk) = (split /class="controlPhoneHide">/, $html);
			 ($phone, $html3) = (split /<\/span>/, $phonejunk);	
			 
	        		$worksheet->write($row, 0, $name);
					$row++;
					$worksheet->write($row, 0, $location);
					$row++;			
					# $worksheet->write($row, 0, $city);
					# $row++;
					# $worksheet->write($row, 0, $state);
					# $row++;
					# $worksheet->write($row, 0, $zip);
					# $row++;
					$worksheet->write($row, 0, $phone);
					$row++;
					$row++;


	
			$stuff_i_want = "";  # important, or you'll keep adding companies each time.
			 
			 
		}
		
		if ($slurping)
		{
		    $stuff_i_want .= $pageline, "\n";
		}
	}
	$pagecount++;
}

$workbook->close();


email({
    'from'         => 'XXXXXXXXX',
    'to'           => 'XXXXXXXXX',
    'smtp'         => 'XXXXXXXXX',
    'subject'      => "Possible Sales Leads",
    '_text'        => "",
    '_attachments' => {
        "$filename" => {
            'description' => "$filename",
            'ctype'       => 'application/vnd.ms-excel',
            'file'        => "$filename",
        },
    },
}) or die "email() failed: $@";
unlink $filename;




## below may be needed depending on how deep we want to dig





















	# $html = $topchunk;
		# ($html1, $namejunk) = (split /<h1 style="margin:9px 5px 0">\n\t\t\t\t/, $html);
		
	
	
	
	# <h1 style="margin:9px 5px 0">
	
	
	
	
	
    # next unless $line =~ /<li><a href="/;     #skip it if it doesnt have the required opening href
	# ($hrefopen, $line2) = (split /<li><a href="/, $line);   # set $line2 to the new link we want plus some closing href crap
	
	# ($line, $hrefclose) = (split /" class="cityCatLink/, $line2);   #set $line to the new link we wanted without closing href crap
	# $lineaddwww = "http://www.partypop.com/$line";   # add the http stuff before the link we liked and set it to $lineaddwww -- this is now the category web address
	
	
	
	# $url = $lineaddwww;       #set $url to the polished link we liked
	# $html = get("$url");      #open the view source for the  link we liked and set it to html -- - $html is now the category url
	
	# ($crap, $crap1) = (split /\?c=/, $html); 
	# ($pagenolist, $crap) = (split /<table class="tBorder1" cellspacing="1"><tr><td class="styleBg1">/, $html1);
	# @pagenolines = split /\n/, $pagenolist;   #separate each line of the page list into @pagenolines
	# foreach $pagenoline (@pagenolines)   # for every line in the pagenumber list, this will eventually set $pagenoline to the web address for another category page
	# {
	    # next unless $line =~ /<a href="/;     #skip it if it doesnt have the required opening href
		# ($pagehrefopen, $pagenoline2) = (split /<a href="/, $pagenoline);   # set $pagenoline2 to the new link we want plus some closing href crap
		# ($pagenoline, $hrefclose) = (split /" class="pageNofM">2<\/a>/, $pagenoline2);   #set $pagenoline to the new link we wanted without closing href crap
		
		# $pageaddress = "http://www.partypop.com$pagenoline";
		# $url = $pageaddress;
		# print $pageaddress;
		# foreach ($pageaddress)
		# {
	        # $html = get("$url");
	        # ($html1, $html2) = (split /invalidPhoneSpan.style.display = 'none';/, $html);   #get rid of crap before new sub category list starts
            # ($subcategories, $html3) = (split /<br><span class="pageNofM">Page<\/span> /, $html2);   # get rid of crap after new sub category list ends
	
	        # @linessec2 = split /\n/, $subcategories;    #separate each line of the new sub category list into @linessec2
	        # foreach $linesec2 (@linessec2)    # for every line in the sub category list
	        # {
	            # next unless $linesec2 =~ /<a href="\/Vendors/;
	            # ($hrefopen2, $linesec22) = (split /<a href="/, $linesec2);
	            # ($linesec2, $hrefclose2) = (split /">/, $linesec22);
	            # $lineaddwww2 = "http://www.partypop.com$linesec2";
	            # $url = $lineaddwww2;
	            # print $url, "\n";
	            # $html = get("$url");
	            # ($html1, $html2) = (split /<img src="http:\/\/www.partypop.com\/Images\/Vendors\//, $html);   #get rid of crap before the name starts
	            # ($namesjunk, $html3) = (split /" align=left>/, $html2);
	            # ($junk, $names) = (split / alt="/, $namesjunk);
		
	            # $html = get("$url");
	            # ($html1, $cityzipjunk, $html2) = (split /<tr><td>&nbsp;<\/td><td>/, $html);   #get rid of crap before the city and zip starts
	            # ($cityzip, $html3) = (split /<\/td><\/tr>/, $cityzipjunk);
		
	            # $html = get("$url");
	            # ($html1, $html2) = (split /<tr><td align=right >Telephone:<\/td><td nowrap >/, $html);   #get rid of crap before the phone number starts
	            # ($phonejunk, $html1, $html3) = (split /<tr><td align=right >Email:/, $html2);
	            # ($phone, $html2) = (split /<\/td><\/tr>/, $phonejunk);
		
		
	            # $html = get("$url");
	            # ($html1, $html2) = (split /<tr><td align=right valign=top >Address:<\/td><td>/, $html);   #get rid of crap before the street starts
	            # ($streetjunk, $html1, $html3) = (split /<tr><td>&nbsp;<\/td><td>/s, $html2); #get rid of crap after the street ends
	            # ($street, $html2) = (split /<\/td><\/tr>/, $streetjunk);
	     
		
	            # foreach ($street)
	            # {
		            # $col=0;
		            # $cityzip =~ s/<\/td><\/tr>//;
		            # $cityzip =~ s/<tr><td>&nbsp;<\/td><td>//;
		            # $cityzip =~ s/\n//;
		            # $cityzip =~ s/\r//;
		            # $street =~ s/<\/td><\/tr>//;
		            # $street =~ s/<tr><td>&nbsp;<\/td><td>//;
		            # $street =~ s/\n//;
		            # $street =~ s/\r//;
		            # last if $street eq "";
		            # last if $street =~ "&nbsp;";
		            # last if $street =~ "Website";
		            # last if $street =~ "website";
		            # last if $street =~ "locations";
		            # last if $street =~ "Locations";
		            # last if $street =~ "Serving";
		            # last if $street =~ "Please";
		            # last if $street =~ "Availability";
		            # last if $street =~ "availability";
		            # last if $street =~ "We Travel";
		            # last if $street =~ "we travel";
		            # last if $street =~ "We travel";
		            # last if $street =~ "we Travel";
		            # last if $street =~ "<tr>";
		            # last if $street =~ "<td>";
		            # last if $street =~ "<\\tr>";
		            # last if $street =~ "<\\td>";
		            # $street =~ s/\t//g;			
		            # $street =~ s/\n\n//g;
		           				
		            # $html = get("$url");
		            # ($html1, $html2) = (split /<img src="http:\/\/www.partypop.com\/Images\/Vendors\//, $html);   #get rid of crap before the name starts
		            # ($namesjunk, $html3) = (split /" align=left>/, $html2);
		            # ($junk, $names) = (split / alt="/, $namesjunk);
		            # $worksheet->write($row, $col, $names);
		            # $col++;
		            
		           
		            # $worksheet->write($row, $col, $street);
		            # $col++;
		            # $worksheet->write($row, $col, $cityzip);
		            # $col++;
		            # $worksheet->write($row, $col, $phone);
		            # $row++;
		            # $number++;
		            # if ($number>8)
		            # {
			            # $workbook->close();
			            # exit;
			        # }
			    # }
			# }
		# }
		
	# }
	
	
	
	
	
# }
# $workbook->close();


# email(XXXXXXXXX
# }) or die "email() failed: $@";
# unlink $filename;
