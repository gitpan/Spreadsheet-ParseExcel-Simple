use ExtUtils::MakeMaker;

WriteMakefile(
    'NAME'		=> 'Spreadsheet::ParseExcel::Simple',
    'VERSION_FROM'	=> 'Simple.pm', # finds $VERSION
    'PREREQ_PM'		=> { 
                             'Spreadsheet::ParseExcel' => 0.18,
                             'Test::More'              => 0.01,
                           },
    ($] >= 5.005 ?    
      (ABSTRACT_FROM => 'Simple.pm', # retrieve abstract from module
       AUTHOR     => 'Tony Bowden <kasei@tmtm.com>') : ()),
);
