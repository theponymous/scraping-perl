#!/usr/local/bin/perl  -w

use strict;
use DBI;
use Spreadsheet::Read;
use Spreadsheet::Read qw(rows);
use Spreadsheet::WriteExcel;
use File::Basename;

XXXXXXXXXXXXXXXX

my
(
    $i,
    $sheet,
    @sheet,
    $ref,
    $rows,
    $cols,
    $row,
    @rowdata,
    $partno, 
    $strSQL, 
    $dbh, 
    $sth,
    $workbook,
    $worksheet,
    $headerformat,
    $filename,
    $sno,
    $ponumber,
    $externalid,
    $desc,
    $qty,
    $atlstock,
    $livstock,
    $cost,
    $eta,
    $location,
    $po_date,
    $asin,
    $seller,
    $price,
    $shipping
);

$::db_connect = $::db_connect; 
$::db_username = $::db_username; 
$::db_password = $::db_password; 

$ref = ReadData("az_sellers.xls");
@sheet = Spreadsheet::Read::rows($ref->[1]);
$rows = $ref->[1]{'maxrow'}; # $rows now equals the total number of rows in the sheet
$cols = $ref->[1]{'maxcol'}; # $cols now equals the total number of columns in the sheet

$filename="sellers_asins.xls";

$workbook = Spreadsheet::WriteExcel->new("$filename");
$worksheet = $workbook->add_worksheet();

$headerformat = $workbook->add_format();
$headerformat->set_bold();
$headerformat->set_underline();
$row=0;

@rowdata = ("Asin", "Partno", "Seller", "Price", "Shipping");
$worksheet->write_row($row, 0, \@rowdata, $headerformat);
$row++;

$dbh = DBI->connect($::db_connect, $::db_username, $::db_password,
                        { PrintError=>0, RaiseError=>0 }) or die "Could not connect to database; Error $DBI::errstr";
$dbh->do("set names utf8");

for ($i=1; $i<$rows; $i++)
{
    $partno=$sheet[$i][0];
    $seller=$sheet[$i][1];
    $price=$sheet[$i][2];
    $shipping=$sheet[$i][3];
    
    next unless $partno;

    $strSQL = <<END;
      select ip_propvalue from inv_properties where ip_inv_partno = '$partno' and ip_propname='amazon_asin'
END
    $sth = $dbh->prepare($strSQL) or die "Could not prepare statment $strSQL; Error $DBI::errstr";
    $sth->execute() or die "Could not execute statement $strSQL; Error $DBI::errstr";
    $sth->bind_columns(undef, \$asin) or die "Could not bind columns for statement; Error $DBI::errstr";
    $sth->fetch();
    $sth->finish();    
        
    $worksheet->write($row, 0, $asin);
    $worksheet->write($row, 1, $partno);
    $worksheet->write($row, 2, $seller);
    $worksheet->write($row, 3, $price);
    $worksheet->write($row, 4, $shipping);
    $row++;
}

$dbh->disconnect() or die "Could not disconnect from database; Error $DBI::errstr";
$workbook->close();
