#! /usr/bin/awk -f
# CSV file structure
# KDNR###KDANREDE###EMAIL###KDNAME1###KDSTRASSE###KDPLZ###KDORT###KDUSTID###STEUERNUMMER###LKZ###DPPROVISION###BANKNAME###BLZ###KTO###PHOTEL###AGTBST

BEGIN { 
	FS="###"
	agency=1
	terminal=16 
}
{
	if ($terminal ~ /,+/) {
		count = split($terminal, sub_array, ",")
		for(i=1; i<=count; ++i) {
			sub_item = sub_array[i]
			if (sub_item in array) {
				repeated_array[sub_item] = sub_item
			} else {
				array[sub_item] = sub_item
			}
		}
	} else if ($terminal in array) {
		repeated_array[$terminal] = $terminal
	} else {
		array[$terminal] = $terminal
	}

}
END {
	for(item in repeated_array) {
		print item
		++n
	}
	print "found: " n
}