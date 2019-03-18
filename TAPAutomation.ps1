#Build the inbox ingestion
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 
# Create a new Comobject which leverages the advantages of the COM interfaces for system administration
$outlook = new-object -comobject outlook.application
# Use the Microsoft Application Programming Interface
$namespace = $outlook.GetNameSpace("MAPI")

$folder = $namespace.getDefaultFolder($olFolders::olFolderInBox)

$emails = $folder.items | Select-Object Body | Select-Object -f 10 
# Build the empty Array to store url links
$URLArray =@()
# loop through all the emails within the inbox
foreach ($email in $emails) {
	#store a matched regex which is a url nd select the url from the stored hash table of $matches,
	# The values is a member of a method from the .NET framework
	$LinksEmail = $email -match"\b(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])"
	$values = $matches | select values
	
	# This is our first inner loop within the loop of email enumeration, at each email within the all emails loop this loop
	# will execute and store the values from the url matches hash table into a position within the
	# $URLArray array data structure
	foreach ($value in $values){$URLArray += $value.values}
	#write-output $value.values
	}
	
	#this is not an inner loop but aloop after we have built our $URLArray array which uses a try-catch 
	# block to attempt to invoke a web request which should be a stored url at each indexed position
	# in the array
	#write-output $URLArray
	foreach ($item in $URLArray) {
	
		try {
			Invoke-WebRequest -verbose $Item
			write-output "This was successful"
			
		}catch { write-output "This Failed $item"}
	}
