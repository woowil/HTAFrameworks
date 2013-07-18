#!/usr/bin/perl -w
# ======================================================================
#
#use IO::Interface;
#use IO::Socket::INET ;
	
sub PLSBtnHello_onClick
{
	#$MyForm->Text1->{'value'} =  "PerlScript says: Hello, world!";
	$window->document->MyForm->Text1->{'value'} =  "PerlScript says: Hello, world!";
}
sub PLSBtnHello_onclick
{
	PLSBtnHello_onClick
}

sub ipconfig {
	
	
	# must create socket first, bound to all interfaces
	$sock = IO::Socket::INET->new(Proto => 'udp');
	
	# get list of active interfaces
	@iflist = $sock->if_list;
	print "List interfaces : @iflist\n";
	
	# get list of active interfaces/addresses
	 for $intf (@iflist) 
	    {
	    $addr = $sock->if_addr(${intf});
	    # creating hash interface/address
	        $addrlist{"$intf"} = "$addr";
	}
	
	# all active interface/ipaddress 
	# using 'key' => 'value' from created hash
	while (($key,$value) = each %addrlist ) {
	  print "$key\t=>\t$value\n"
	}
}
