use strict ;
use warnings ;

use Excel::Writer::XLSX ;
use Spreadsheet::Read qw ( ReadData ) ;
use DateTime::Format::Excel ;
use Data::Dumper ;
print ( "Hello !\nPlease type/paste in the full name of the directory you wish to submit\n" ) ;
my $dir = <STDIN> ;
chomp $dir ;
# Substitue all \ for / for Perl to read correctly
$dir =~ s/\\/\//g ;
my $date_format = "m/d/yy" ;
my $output_cell_val ;
my $date_conversion ;

my @temp_arr ;
my $output_xlsx = Excel::Writer::XLSX->new( "output_xlsx.xlsx" ) ;
my $output_sheet = $output_xlsx->add_worksheet() ;

my $num_files = 0 ;
while ( $_ = glob( "$dir/*.xlsx" ) )
{
	print ( "$_\n" ) ; 
	my @temp_arr ;
	my $curr_wb = ReadData( "$_", parser => "xlsx" , dtfmt => "$date_format" ) ;
	my $sheet = $curr_wb->[3]{label};
	for my $index ( 1..26 )
	{
		$output_cell_val = $curr_wb->[3]{cell}[$index][18] ;

		if ( ( $output_cell_val ne "" ) && ( $index != 1 ) )
		{
			if ( $output_cell_val eq "0" )
			{
				$output_cell_val = "" ;
			}
			# write to output excel
			push ( @temp_arr , "$output_cell_val" ) ;

		}
			
		if ( $index == 1 )
		{
			# excel stores dates as numeric values starting at Jan 1, 1900 as 0
			$date_conversion = DateTime::Format::Excel->parse_datetime( $output_cell_val ) ;
			my $output_date = $date_conversion->mdy( '/' ) ;
			# write to output excel
			push ( @temp_arr , "$output_date" ) ;
		}
	}
	# write_row uses array reference as input rather than array
	my $temp_arr_ref = \@temp_arr ;
	$output_sheet->write_row ( $num_files, 0, $temp_arr_ref ) ;
	@temp_arr = () ;
	$num_files++ ;
	# next line in output excel
}