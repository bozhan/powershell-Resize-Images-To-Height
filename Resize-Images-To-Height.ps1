 param (
    [string]$d,
    [String[]]$t = ('.jpg'),
		[string]$h = 2160,
    [switch]$r,
		[switch]$help
 )

$Missing_Filetype_Error = New-Object System.FormatException "-t (file type) is missing!"
$Missing_Folder_Error = New-Object System.FormatException "-d (source folder) is missing!"

filter Where-Extension {
param([String[]] $extension = $t)
	$_ | Where-Object {
		$extension -contains $_.Extension
	}
}

function Get-Script-Folder-Path{
  return Split-Path -Path $script:MyInvocation.MyCommand.Path -Parent
}

function Get-ImageMagick-Converter-Path{
	$scriptFolder = Get-Script-Folder-Path
	$execName = "imgMagickConvert.exe"
	$filesInScriptFolder = @(Get-ChildItem -Recurse -literalPath $scriptFolder)
	$converter = ""
	if ($filesInScriptFolder.count -gt 0){
		foreach($file in $filesInScriptFolder){
			if ($file.Name.ToLower().CompareTo($execName.ToLower()) -eq 0){
				$converter = $file.FullName
			}
		}
	}
	if (($filesInScriptFolder.count -eq 0) -or ($converter.CompareTo("") -eq 0)){
		Throw [System.IO.FileNotFoundException] "Executable ImageMagick convert.exe missing from folder:""$scriptFolder"""
	}
	return $converter
}

function Get-Height($f){
	$shellObject = New-Object -ComObject Shell.Application
  $heightAttribute = 164
	$directoryObject = $shellObject.NameSpace( $f.Directory.FullName )
	$fileObject = $directoryObject.ParseName( $f.Name )
			
	# Find the index of the bit rate attribute, if necessary.
	for( $index = 5; -not $heightAttribute; ++$index ) {
		$name = $directoryObject.GetDetailsOf( $directoryObject.Items, $index )
		if( $name -eq 'Height' ) {
			$heightAttribute = $index
		}
	}
	
	# Get the bit rate of the file.
	$heightString = $directoryObject.GetDetailsOf( $fileObject, $heightAttribute )
	if( $heightString -match '\d+' ) { 
		[int]$height = $matches[0] 
	}else { 
		$height = -1 
	}
	
	return $height
}

function Resize-To-Height([string]$folderPath = $d, [int]$desiredHeight = $h, [boolean]$recurse = $r){
	$convertPath = Get-ImageMagick-Converter-Path
	
	if($recurse){
		$files = @(Get-ChildItem -literalPath $folderPath -recurse | Where-Extension $t)
	}else{
		$files = @(Get-ChildItem -literalPath $folderPath | Where-Extension $t)
	}
	
	write-host "File count: " $files.Count
	if($files.count -gt 0){
		foreach($file in $files){
			$i = $i + 1
			$height = Get-Height $file
			if ($height -gt $desiredHeight){
				write-host $i "/" $files.count " converted " $file.Name " with " $height "to" $h
        &$convertPath $file.FullName -resize x$h $file.FullName | Out-Null
			}
		}
	}
}

if($help.ispresent){
	write-host("`nSYNTAX")
	write-host("  " + "Resize-Images-To-Height [[-d]<string>] [[-t]<string[]=('.jpg')>] [[-h]<integer=2160>] [[-r]<switch>]")
	write-host("`nDESCRIPTION")
	write-host("  " + "{0,-15} {1}" -f "-d", "Specify directrory to containing media files.")
	write-host("  " + "{0,-15} {1}" -f "-t", "Specify media file types.")
	write-host("  " + "{0,-15} {1}" -f "-h", "Set maximum height, all images with height heigher than that value will be donwsampled to it.")
	write-host("  " + "{0,-15} {1}" -f "-r", "Search directory recursively for files.")
	exit
}

$d = $(Read-Host 'Source Folder')
if (-not $d){throw $Missing_Folder_Error}

if ($r){
	$rec = $true
}else{
	$rec = $false
}

Resize-To-Height $d $h $rec