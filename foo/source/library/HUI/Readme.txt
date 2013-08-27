Copyright (c) 2003-2013 nOsliw HUI - HTML/HTA Application Framework, All Rights Reserved

Name             : nOsliw HUI - HTML/HTA Application Framework
Copyright holder : W. Wilson
Company          : nOsliw Solutions
Open source site : http://hui.codeplex.com
License          : http://hui.codeplex.com/license
Releases         : http://hui.codeplex.com/releases
Donate           : https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=5805600

This library is under the terms of the "GNU Library General Public License (LGPL)" license
Please remember that this is an early version of the library, not all features are implemented yet, and not everything might work as expected, including documentation. Visit the open sourse site for bug reports, feature requests, development or general dicussion.


Launch from HTA
....

<HEAD>
<SCRIPT language="JScript.Encode" src="source/library/HUI/Base/HUILoader.js" id="JSLoad-0" charset="UTF-8" type="text/javascript"></script>
<script language="JScript">

function mba_onload(){

	HUILoader()
	
	if(!__H.initialize({
		textarea_result : oFormDebug.t_log,
		textarea_error  : oFormDebug.t_error,
		file_log		: "debug.log",
		file_error		: "debug-error.log",
		include_plugins : true,
		include_all     : true
	})) return;
}

</HEAD>	
......



