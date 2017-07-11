#!/usr/bin/perl -w

use Text::Iconv;
my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
 
use Spreadsheet::XLSX;
use Excel::Writer::XLSX;
 
# Excel sheet is a single column, list of strings (funds to exclude)
my $exclude        = Spreadsheet::XLSX -> new ('exclude.xlsx', $converter);
my @exclude_sheets = @{$exclude->{Worksheet}};
my $exclude_sheet  = $exclude_sheets[0];

my @exclude_list = ();

foreach my $row ($exclude_sheet->{MinRow} .. $exclude_sheet->{MaxRow}) {
  if ($exclude_sheet->{Cells}[$row][0]->{Val}) {
    push(@exclude_list, $exclude_sheet->{Cells}[$row][0]->{Val});
  }
}

my $directory=".";
opendir(DIR, $directory) or die "couldn't open $directory: $!\n";
my @files = readdir DIR;
closedir DIR;

foreach my $file (@files) {
  if ($file =~ /xlsx/ && $file !~ /exclude/) {
    my $input     = Spreadsheet::XLSX -> new ($file, $converter);
    my @sheets    = @{$input->{Worksheet}};
    my $in_sheet  = $sheets[0];

    my $output    = Excel::Writer::XLSX->new('out-' . $file);
    my $out_sheet = $output->add_worksheet();

    my $out_row = 0;
    foreach my $row ($in_sheet->{MinRow} .. $in_sheet->{MaxRow}) {
 
      my $txn_type = $in_sheet->{Cells}[$row][0]->{Val};
      if ($txn_type && ($txn_type =~ /sl/ || $txn_type =~ /by/)) {

        my $trade_date = $in_sheet->{Cells}[$row][2]->{Val};
 
        # trade dates between the 28th and 6th, inclusive
        if ($trade_date =~ /\d+-\d+-(\d+)/) {
          my $day = $1;
          if ($day >= 28 || $day <= 6) {

            my $security = $in_sheet->{Cells}[$row][1]->{Val};
            my $matched_exclusion = 0;
            foreach my $exclusion (@exclude_list) {
              if ($security =~ /$exclusion/i) {
                $matched_exclusion = 1;
                last;
              }
            }

            if (!$matched_exclusion) {
              print "Writing row $out_row\n";
              # output data
              for my $col (0 .. 9) {
                if ($in_sheet->{Cells}[$row][$col]) {
                  $out_sheet->write($out_row, $col, $in_sheet->{Cells}[$row][$col]->{Val});
                }
              }
              $out_row++;
            } else {
              print "Matched entry from exclusion list. Skipping.\n";
            }

          }
        }

      }

    }
  }
}
