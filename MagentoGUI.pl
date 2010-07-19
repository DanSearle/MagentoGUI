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
use SOAP::Lite ; #+trace => 'debug';

## TK GUI library
use Tk; 
use Tk::FileSelect;
use Tk::ProgressBar;

## CSV Parser library
use Text::CSV;

## Base64 encoder library
use MIME::Base64;

## Config
#my $SOAPURL = "http://dan.homelinux.net/magento-testing/magento/index.php/api/soap/?wsdl";
#my $USER    = "test";
#my $PASS    = "test123";
my $SOAPURL = "http://dan.homelinux.net/magento/index.php/api/soap/?wsdl";
my $USER    = 'mike';
my $PASS    = 't7x93kaMKR2THJjBswQUiVMizCTGEjEkeRcR75o8jmuuk0QAk79p7uDEXGgGPTVW';
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
startGUI();

# Start the interface
sub startGUI {
    $MainWindow = MainWindow->new(); # Create a main window
    $MainWindow->geometry('600x300+0+0');

    # Create the browse file frame
    my $filefrm = $MainWindow->Frame()->pack(-fill => "x", -expand => 1);
    # Add the label, textbox and browse button
    $filefrm->Label (-text         => 'Filename:'    )->pack(-side => 'left');
    $filefrm->Button(-text         => 'Browse...', 
                     -command      => \&browseFileGUI)->pack(-side => 'right');
    $filefrm->Entry (-textvariable => \$Filename     )->pack(-side => 'right', -fill => "x", -expand => 1);

    # Create the upload button
    $MainWindow->Button(-text => 'Upload', -command => \&upload)->pack();

    # Create a progress bars
    $CheckProgressBar  = $MainWindow->ProgressBar(-foreground => 'red')->pack(-fill => "x", -expand => 1);
    $UploadProgressBar = $MainWindow->ProgressBar(-foreground => 'green')->pack(-fill => "x", -expand => 1);

    # Create the scrolling logging window
    my $txt = $MainWindow->Text()->pack(-fill => "both", -expand => 1);
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
    die("Cannot call the API if we are not connected to the SOAP server") unless ($Server || $SessionID);
    my @args = ($SessionID, @_);
    return $Server->call(@args) or die("Could not execute API call!");
}

# Login to the server
sub login {
    $Server    = SOAP::Lite
                    ->service($SOAPURL)
                    #->use_prefix(1)
                    ->envprefix('SOAP-ENV')
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
  open my $CSVFile, $filename or (logMessage("$!") and return);

  # Create a csv object
  my $csv = Text::CSV->new({binary => 1, eol => $/ });

  # Read the file line by line
  while(my $values = $csv->getline($CSVFile)) {
    next unless ($. > 1); # Skip the first line

    # Replace newlines with html line breaks
    $values->[2] =~ s/\n/<br\/>/g;
    $values->[3] =~ s/\n/<br\/>/g;

    # Get the values
    my $data = {
                  SKU               => $values->[0],
                  Name              => $values->[1],
                  Short_Description => $values->[2],
                  Description       => $values->[3],
                  Price             => $values->[4],
                  Attribute_Set     => $values->[5],
                  Attributes        => 
                  [map {my @s = split(/:/, $_);
                                           my $name = $s[0];
                                           my $val  = $s[1];
                                           { 'name' => $name, 'value' => $val }
                                          } split (/;/, $values->[6])],
                  Image             => $values->[7],
                  Url_Key           => $values->[8],
                  Related           => [split(/;/, $values->[9])],
                  Cat_ID            => $values->[10]
              };

    # Add the data to the list of records
    push (@records, $data);
  }
  $csv->eof or (logMessage("Failed to parse line: " . $csv->error_input . "\n" . "Message: " . Dumper($csv->error_diag())) and return -1);

  # Close the file
  close $CSVFile;

  return @records;
}

# Start the upload process
sub upload {
    die("No filename provided, cannot do anything.") unless ($Filename);
    # Clear progress bars
    $CheckProgressBar->value(0);
    $UploadProgressBar->value(0);
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
        my $retattr;
        eval {
            $retattr = callAPI('product.info', [$_->{SKU}]);
        };
        chomp($@);
        die("Product " . $_->{SKU} . " already exists!") unless (($@ eq 'Product not exists') or ($retattr->{'sku'} ne $_->{SKU}));
        
        # Check for non existent related products
        # TODO: Cache for speed
        if($_->{Related} > 0) {
            foreach my $relsku ($_->{Related}) {
                my $relretattr; 
                eval {
                   $relretattr = callAPI('product.info', [$relsku]);
                };
                chomp($@);
                die("Related product: " . $$relsku[0] . " for product " . $_->{SKU} . " does not exist!") unless (($@ ne 'Product not exists.') or ($relretattr->{'sku'} eq $relsku));


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
          die("Attribute " . $attr->{name} . " does not exist in the attribute set...") unless($_->{Attribute_Set_Name_ID}->{$attr->{name}});
          # Get the attribute options
          my @options = @{callAPI('product_attribute.options', [$attr->{name}])};

          foreach my $option (@options) {
            if ($option->{label} eq $attr->{value}) {
              # Set the attribute in the product data
              $productData->{$attr->{name}} = $option->{value};
            }
          }
        }

        # Add the basic details of the product
        $productData->{'name'} = $_->{Name};
        $productData->{'short_description'} = $_->{Short_Description};
        $productData->{'description'} = $_->{Description};
        $productData->{'price'} = $_->{Price};
        $productData->{'url_path'} = $_->{Url_Key} . ".html";
        $productData->{'url_key'} = $_->{Url_Key};
        $productData->{'status'} = 1;
        $productData->{'weight'} = 0;
        $productData->{'tax_class_id'} = 2;
        $productData->{'websites'} = [1];
        $productData->{'stock_data'} = SOAP::Data->type('map' => { 
                                         'min_sale_qty'            => 1,
                                         'use_config_min_sale_qty' => 1,
                                         'use_config_max_sale_qty' => 1,
                                         'use_config_manage_stock' => 1
                                       });


        #my $productReq = ['simple', $_->{Attribute_Set_ID} , $_->{SKU}, $productData];
        my $productReq = ['simple', $_->{Attribute_Set_ID}, $_->{SKU}, SOAP::Data->type('map' => $productData)];

        # Actually create the product
        callAPI('catalog_product.create', $productReq) or die("Product creation failed");
        
        logMessage("Adding related products");
        if(@{$_->{Related}}) {
          foreach my $rel (@{$_->{Related}}) {
            callAPI('product_link.assign', ['related', $_->{SKU}, $rel]);
          }
        }
        #FIXME Needs testing

        # Upload the image if any
        if($_->{Image} ne "") {
          open(IMG, $_->{Image}) or die("Could not open image file!");
          local($/) = undef; # slurp
          my $base64content = MIME::Base64::encode(<IMG>);
          close(IMG);
          callAPI('product_media.create', [$_->{SKU}, SOAP::Data->type( "map" => {
                                                       file => SOAP::Data->type("map" => { 
                                                         content => $base64content, 
                                                         mime => "image/jpeg"
                                                       }), 
                                                       types => ["small_image", 
                                                                 "image", 
                                                                 "thumbnail"], 
                                                       exclude => 0
                                                     })]);
        }
  
        # Add the product to a category
        logMessage("Adding the product to its category");
        if($_->{Cat_ID} ne "") {
          callAPI('category.assignProduct', [$_->{Cat_ID}, $_->{SKU}]);
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
#$Filename = "/home/dan/Dropbox/Programming/MagentoGUI/testing.csv";
#startGUI();
