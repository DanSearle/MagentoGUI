#!/usr/bin/perl -w
# MagentoGUI - Graphical program to upload products to a magento install
#              from a spreadsheet
#
#   Author: Daniel Searle
#
# vim:ts=4 sw=4 autoindent expandtab
#

## Helpful perl warnings and errors
use warnings;
use strict;
use Data::Dumper;

## SOAP api library
use SOAP::Lite +trace => 'debug';

## TK GUI library
use Tk; 
use Tk::FileSelect;
use Tk::ProgressBar;

## CSV Parser library
use Text::CSV;

## Config
my $SOAPURL = "http://dan.homelinux.net/magento-testing/magento/index.php/api/soap/?wsdl";
my $USER    = "test";
my $PASS    = "test123";
## End Config

## Globals
my $Server;
my $SessionID;
my $MainWindow;
my $Filename;
my $CheckProgressBar;
my $UploadProgressBar;
## End Globals

# Launch the GUI
#startGUI();

# Start the interface
sub startGUI {
    $MainWindow = MainWindow->new(); # Create a main window
    $MainWindow->geometry('600x300+0+0');

    # Create the browse file frame
    my $filefrm = $MainWindow->Frame()->pack();
    # Add the label, textbox and browse button
    $filefrm->Label (-text         => 'Filename:'    )->pack(-side => 'left' );
    $filefrm->Button(-text         => 'Browse...', 
                     -command      => \&browseFileGUI)->pack(-side => 'right');
    $filefrm->Entry (-textvariable => \$Filename     )->pack(-side => 'right');

    # Create the upload button
    $MainWindow->Button(-text => 'Upload', -command => \&upload)->pack();

    # Create a progress bars
    $CheckProgressBar  = $MainWindow->ProgressBar(-length => 600, -foreground => 'red')->pack();
    $UploadProgressBar = $MainWindow->ProgressBar(-length => 600, -foreground => 'green')->pack();

    # Create the scrolling logging window
    my $txt = $MainWindow->Text(-width=>85, -height=>17)->pack();
    tie *STDOUT, ref $txt, $txt;
    tie *STDERR, ref $txt, $txt;

    MainLoop();
}

# Launch the browse dialog to set the filename
sub browseFileGUI {
    my $fsref = $MainWindow->FileSelect();
    $Filename = $fsref->Show;
}

# Log a message
sub logMessage {
    print shift() . "\n";
    $MainWindow->update();
}

# Call the API 
# 
# Args:
#      Array - Call arguments, without the sessionID
sub callAPI {
    die("Cannot call the API if we are not connected to the SOAP server") if (!$Server || !$SessionID);
    my @args = ($SessionID, @_);
    return $Server->call(@args) or die("Could not execute API call!");
}

# Login to the server
sub login {
    $Server    = SOAP::Lite
                    ->service($SOAPURL)
                    ->on_fault(\&SOAPon_fault) or die("Cannot connect to SOAP server");

    $SessionID = $Server->login($USER, $PASS)  or die("Failed to login to the SOAP server");
}

# Catch a fault with the soap transaction
sub SOAPon_fault {
    my($soap, $res) = @_; 
    die ref $res ? $res->faultstring : $soap->transport->status, "\n";
}

# Logout of the server
sub logout {
    $Server->endSession($SessionID) or die("Could not end the session");
}

# Parse the CSV file into a array of hashes
# Args:
#   Filename
sub parseCSV {
  # Filename of the file
  my $filename = shift();
  
  # List of hashes representing the records
  my @records;

  # Text of the current line
  my $currline = "";

  # Open the file
  open CSV, $filename or (logMessage("$!") and return);

  # Create a csv object
  my $csv = Text::CSV->new();

  # Read the file line by line
  while(<CSV>) {
    next unless ($. > 1); # Skip the first line
    $currline .= $_;      # Append to the current line

    # If the line has not ended add a html newline
    # and finish parsing the record, by moving to the
    # next line
    ($currline =~ s/\n$/<br\/>/g and next) if (! ($currline =~ /"$/ || $currline =~ /[0-9]$/));
  
    # Parse the line and print a error message and return if it fails.
    # (logMessage("Failed to parse line: " . $csv->error_input ) and return -1) if !(;
    if (! ($csv->parse($currline))) {
      logMessage("Failed to parse line: " . $csv->error_input . "\n" . "Message: " . Dumper($csv->error_diag));
      return -1;
    }

    # Blank the parsed line for a new run
    $currline = "";

    # Get the values
    my @values = $csv->fields();
    my $data = {
                  SKU               => $values[0],
                  Name              => $values[1],
                  Short_Description => $values[2],
                  Description       => $values[3],
                  Price             => $values[4],
                  Attribute_Set     => $values[5],
                  Attributes        => 
                  [map {my @s = split(/:/, $_);
                                           my $name = $s[0];
                                           my $val  = $s[1];
                                           { 'name' => $name, 'value' => $val }
                                          } split (/;/, $values[6])],
                  Image             => $values[7],
                  Url_Key           => $values[8],
                  Related           => [split(/;/, $values[9])],
                  Cat_ID            => $values[10]
              };

    # Add the data to the list of records
    push (@records, $data);
  }

  # Close the file
  close CSV;

  return @records;
}

# Start the upload process
sub upload {
    # Login
    logMessage("Logging into Magento...");
    login();

    # Parse the CSV file
    my @csv = parseCSV($Filename);

    # Cache the SKUs
    my @skus;
    # Cache the attribute sets
    my @attributeSetsAPI  = @{callAPI('product_attribute_set.list')};
    my %attributeSets;
    foreach (@attributeSetsAPI) {
        $attributeSets{$_->{name}} = $_->{set_id};
    }
    my @attributeSetNames;
    foreach(@attributeSetsAPI) {
        push(@attributeSetNames, $_->{name});
    }

    # Check for duplicate items
    logMessage("Checking sanity of spreadsheet...");
    my $progress = 0;
    foreach(@csv) {
        # See if the product already exists
        eval {
            callAPI('product.info', [$_->{SKU}]);
        };
        chomp($@);
        die("Product " . $_->{SKU} . " already exists!") if ($@ ne 'Product not exists.');
        
        # Check for non existent related products
        # TODO: Cache for speed
        if($_->{Related} > 0) {
            foreach my $relsku ($_->{Related}) {
                eval {
                    callAPI('product.info', [$relsku]);
                };
                chomp($@);
                die("Related product: " . $relsku . " for product " . $_->{SKU} . " does not exist!") if ($@ eq 'Product not exists.');

                $MainWindow->update();
            }
        }
    
        # Check to see if the image file can be read
        die("Image file " . $_->{Image} . " does not exist, or cannot be read for SKU: " . $_->{SKU}) unless ((-r $_->{Image} && -f $_->{Image}) || ($_->{Image} eq ""));

        # Check to see if the attribute set exists
        my $attrset = $_->{Attribute_Set};
        die("Attribute set " . $_->{Attribute_Set} . " does not exist for SKU: " . $_->{SKU}) unless ($attributeSets{$attrset});

        # Cache the SKUs
        push(@skus, $_->{SKU});
        
        # Update the progress bar
        $progress++;
        $CheckProgressBar->value(($progress/@csv)*100);

        # Update the mainwindow to stop freezing
        $MainWindow->update();
    }

    # Check for duplicates within the spreadsheet
    logMessage("Checking for duplicate SKU's in the spreadsheet...");
    foreach my $sku (@skus) {
            die("SKU: " . $sku . " exists more than once in the spreadsheet") unless (grep(/^$sku$/, @skus) == 1);
    }

    # Do actual upload
    my $productData;
    logMessage("Pre-flight checks passed, proceeding with upload...");
    $progress = 0;
    foreach (@csv) {
        # Upload the product info
        logMessage("Uploading product: " . $_->{SKU});
        $_->{Attribute_Set_ID} = $attributeSets{$_->{Attribute_Set}};
        $_->{Attribute_Set_List} = callAPI('product_attribute.list', [$_->{Attribute_Set_ID}]);
        $_->{Attribute_Set_Name_ID} = {map {my $name = $_->{'code'};
                                            my $id   = $_->{'attribute_id'};
                                            $name => $id
                                           } @{$_->{Attribute_Set_List}}};

        # Add the attributes to the product data
        foreach my $attr (@{$_->{Attributes}}) {
          die("Attribute " . $attr->{value} . " does not exist in the attribute set...") unless($_->{Attribute_Set_Name_ID}->{$attr->{name}});
          # Set the attribute in the product data
          $productData->{$attr->{name}} = $attr->{value};
        }

        # Add the basic details of the product
#        $productData->{'name'} = $_->{Name};
#        $productData->{'short_description'} = $_->{Short_Description};
#        $productData->{'description'} = $_->{Description};
#        $productData->{'price'} = $_->{Price};
#        $productData->{'url_path'} = $_->{Url_Key} . ".html";
#        $productData->{'url_key'} = $_->{Url_Key};
#        $productData->{'status'} = 1;
#        $productData->{'weight'} = 0;
#        $productData->{'tax_class_id'} = 2;
#        $productData->{'websites'} = [1];
#        $productData->{'stock_data'} = { 'min_sale_qty'            => 1,
#                                         'use_config_min_sale_qty' => 1,
#                                         'use_config_max_sale_qty' => 1,
#                                         'use_config_manage_stock' => 1
#                                       };

        print Dumper $_->{Attribute_Set_ID};
        print Dumper $_->{SKU};
        print Dumper $productData;

        #my $productReq = ['simple', $_->{Attribute_Set_ID} , $_->{SKU}, $productData];
        my $productReq = ['simple', $_->{Attribute_Set_ID}, $_->{SKU}, $productData];

        # Actually create the product
        callAPI('catalog_product.create', $productReq) or die("Product creation failed");
        
        #FIXME Needs testing
        foreach my $rel ($_->{Related}) {
          callAPI('product_link.assign', ['related', $_->{SKU}, $rel]);
        }

        # Upload the image if any
        if($_->{Image} ne "") {
          open(IMG, $_->{Image}) or die("Could not open image file!");
          local($/) = undef; # slurp
          my $base64content = MIME::Base64::encode(<IMG>);
          close(IMG);
          callAPI('product_media.create', [$_->{SKU}, {file => { 
                                                       content => $base64content, 
                                                       mime => "image/jpeg"}, 
                                                     types => ["small_image", 
                                                               "image", 
                                                               "thumbnail"], 
                                                     exclude => 0}]);
        }
  
        # Add the product to a category
        if($_->{Category} ne "") {
          callAPI('category.assignProduct', [$_->{Category}, $_->{SKU}]);
        }

        # Update the progress bar
        $progress++;
        $UploadProgressBar->value(($progress/@csv)*100);
        $MainWindow->update;
    }
    
    # Logout
    logMessage("Logging out of Magento...");
    logout();
}

# TESTING
#login();
$Filename = "/home/dan/Dropbox/Programming/MagentoGUI/testing.csv";
startGUI();
