#!/usr/local/bin/perl  -w

use strict;
use DBI;
use LWP::Simple;
use LWP::UserAgent;
use Spreadsheet::Read;
use Spreadsheet::Read qw(rows);
use Spreadsheet::WriteExcel;
use File::Basename;
use Time::HiRes;


XXXXXXXXXXXX


my (
    $i,
    $sheet,
    @sheet,
    $ref,
    $rows,
    $cols,
    $row,
    $partno, 
    $filename,
    $html,
    $crap,
    $crap1,
    $crap2,
    $crap3,
    $price,
    $good,
    $good1,
    $good2,
    $workbook,
    $worksheet,
    $headerformat,
    @rowdata,
    $rowdata,
    $asin,
    $url,
    @temp,
    @lines,
    $line,
    $vendorid,
    $shipping,
    $cookie_jar_obj,
    $seller,
    $sellerline,
    @sellerlines,
    $sellerurl,
    $cwsprice,
    %sellernames,$sellerprice,
    $dbh,           # database handle
    $sth,           # statement handle
    $strSQL        # database SQL statement
);

# $::db_connect = $::db_connect; 
# $::db_username = $::db_username; 
# $::db_password = $::db_password; 
use WWW::Mechanize;
my $mech = WWW::Mechanize->new( 'ssl_opts' => { 'verify_hostname' => 0,  'SSL_verify_mode' => 'SSL_VERIFY_CLIENT'}, 'onerror' => undef );
$mech->agent_alias( 'Windows Mozilla' );
XXXXXXXXXXX
$mech->get('https://sellercentral.amazon.com/gp/homepage.html' );

# $mech -> form_name('signIn');
# $mech -> field ('email' => 'XXXXXXXX');
# $mech -> field ('password' => 'XXXXXXXXX');
# $mech -> click_button( value => 'Sign in using our secure server ' );

$mech -> form_name('signIn');
$mech -> field ('username' => 'XXXXXXX');
$mech -> field ('password' => 'XXXXXXXX');
$mech -> click_button( name => 'sign-in-button' );

$ref = ReadData("XXXXXXX.xls");
@sheet = Spreadsheet::Read::rows($ref->[1]);
$rows = $ref->[1]{'maxrow'}; # $rows now equals the total number of rows in the sheet
$cols = $ref->[1]{'maxcol'}; # $cols now equals the total number of columns in the sheet

$filename="sellers1.xls";

$workbook = Spreadsheet::WriteExcel->new("$filename");
$worksheet = $workbook->add_worksheet();

$headerformat = $workbook->add_format();
$headerformat->set_bold();
$headerformat->set_underline();
$row=0;

@rowdata = ("Asin", "Partno", "Seller ID", "Seller", "Price", "CWS Price");
$worksheet->write_row($row, 0, \@rowdata, $headerformat);
$row++;

open(OUT, '>', 'XXXXXXXX.txt');


for ($i=1; $i<$rows; $i++)
{
    Time::HiRes::sleep(2.000);  # 1.2 seconds
    $partno=$sheet[$i][0];
    $asin=$sheet[$i][1];
    # $asin="XXXXXXXX";
    $cwsprice=$sheet[$i][2];
    print "Attempting to scrape $partno...\n";
    $url="http://www.amazon.com/gp/offer-listing/$asin/ref=olp_tab_new?ie=UTF8&condition=new";

    $mech->get( $url );
    if ($mech->status() eq "404")
    {
        print "404 on part # $partno\n";
        print OUT "$partno\t$asin\n";
        next;
    }
    
    $html = $mech->content;
    
    @temp = split /\n/, $html;
    $html = join("", @temp);
    # print OUT "$html\n";
    $html =~ s/a-text-center a-spacing-large/a-spacing-mini a-divider-normal/g;
    
    @lines = $html =~ m/a-row a-spacing-mini olpOffer.+?a-spacing-mini a-divider-normal/sg;

    foreach $line(@lines)
    {
        print OUT "$line\n\n----------------------------------------------------\n\n";
        #print "$line\n\n";
        next if ($line =~ m/01dXM-J1oeL.gif/i);
        next if ($line =~ m/Available to buy on/i);
        
        ($crap, $good) = split /seller=/, $line;
        ($vendorid, $crap1) = split /&amp;/, $good;  
        $vendorid =~ s/ //g; 
        
        if ($vendorid eq "XXXXXXXXXX" or 
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX" or
        $vendorid eq "XXXXXXXXXX")
        {
            next;
        }
        
        ($crap, $good) = split /olpOfferPrice a-text-bold">/, $line;
        ($price, $crap1) = split /<\/span>/, $good;
        $price =~ s/^\s*(.*)?/$1/;  # kill leading spaces
        $price =~ s/\s+$//;  # kill trailing spaces
        $price =~ s/\$//g; 
        
        ($crap, $good) = split /<span class="a-color-secondary">/, $line;
        ($shipping, $crap1) = split /<\/span>/, $good;
        $shipping =~ s/^\s*(.*)?/$1/;  # kill leading spaces
        $shipping =~ s/\s+$//;  # kill trailing spaces
        if ($shipping =~ m/Free Shipping/i)
        {
            $shipping = "0";
        }
        else
        {
            ($crap, $good) = split /olpShippingPrice">/, $line;
            ($shipping, $crap1) = split /<\/span>/, $good;
            $shipping =~ s/^\s*(.*)?/$1/;  # kill leading spaces
            $shipping =~ s/\s+$//;  # kill trailing spaces
            $shipping =~ s/\$//g;  
        }
        
        $sellerprice = $price+$shipping;
        # next and print "skipping $vendorid" unless ($sellerprice < .65*$cwsprice);
        
        $seller = GetSellerName($vendorid); 
        
        print "$seller\t$sellerprice\t$cwsprice\n";
         
        $worksheet->write($row, 0, $asin);
        $worksheet->write($row, 1, $partno);
        $worksheet->write($row, 2, $vendorid);
        $worksheet->write($row, 3, $seller);
        $worksheet->write($row, 4, $sellerprice);
        $worksheet->write($row, 5, $cwsprice);
        $row++;

        $seller = "";
        $sellerprice = "";
        $shipping = "";
        $price = "";
    }
}
$workbook->close();

close OUT;

sub GetSellerName
{
  my ($refsellerid) = @_;
  my (
      $sellername,
      $sellerurl,
      $crap,
      $crap1,
      $crap2,
      $crap3,
      $good,
      $good1,
      $good2,
      @sellerlines,
      $sellerline,
      $returnval,
     ) ;
     
  if ($sellernames{$refsellerid}) # read from cache if possible :)   
  {
      print "found $sellernames{$refsellerid} in cache\n";
      $returnval = $sellernames{$refsellerid};
  }
  else
  {
      #dig into sellers store front to get name
      $sellerurl = "http://www.amazon.com/s?merchant=$refsellerid";
      
      $mech->get( $sellerurl );
      if ($mech->status() eq "404")
      {
          print "404 on part # $partno\n";
          return "not found";
      }
      # print OUT $mech->content . "\n\n";
      @sellerlines = split /\n/, $mech->content;
      foreach $sellerline(@sellerlines)
      {
          if ($sellerline =~ m/<span class="nav-search-label">/)
          {
              ($crap, $good) = split /<span class="nav-search-label">/, $sellerline;
              ($sellername, $crap1) = split /<\/span>/, $good;
              $sellername =~ s/^\s*(.*)?/$1/;  # kill leading spaces
              $sellername =~ s/\s+$//;  # kill trailing spaces
          }
          if ($sellerline =~ m/captcha/i)
          {
              print "found captcha\n";
          }
      }

      print "had to scrape to get $sellername\n";

      $sellernames{$refsellerid} = $sellername;
      $returnval = $sellernames{$refsellerid};
  }
      
  return $returnval;
}

