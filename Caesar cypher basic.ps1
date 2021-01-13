###################
#Encrypt
###################

#Message to encrypt
$text = "titi a cru voir un grosminet"

#Shift number
$i = 3

#Transform the text to an array of ascii
$asciis = [int[]][char[]]"$text"

#Add the shift specified to every ascii character
$encrypt = foreach ($ascii in $asciis) {
$ascii + $i
}

#Transform back the ascii to text
$asciitotext = [char[]]$encrypt

#Transform the array to a string to read it easily
$arraytostring = -join $asciitotext
$arraytostring

###################
#Decrypt
###################

$text = "wlwl#d#fux#yrlu#xq#jurvplqhw"

#Shift number
$i = 3

#Transform the text to an array of ascii
$asciis = [int[]][char[]]"$text"

#Add the shift specified to every ascii character
$decrypt = foreach ($ascii in $asciis) {
$ascii - $i
}

#Transform back the ascii to text
$asciitotext = [char[]]$decrypt

#Transform the array to a string to read it easily
$arraytostring = -join $asciitotext
$arraytostring
