<# Written by Joe Gasparich #>

Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"

<# CREDENTIALS #>

function Get-SPCredentials() {
	<#
        .SYNOPSIS
            Returns a credentials object for use with SharePoint CSOM
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.SharePointOnlineCredentials])]
	param (
		# The username of the credentials
		[String]
		$UserName,

		# The password of the credentials
		[System.Security.SecureString]
		$Password
	)
	process {
		return New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $Password)
	}
}


<# CONTEXT #>

function Get-SPContext() {
	<#
        .SYNOPSIS
            Returns a SharePoint client context object from a url and credentials
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ClientContext])]
	param (
		# The URL of the site to get context of
		[String]
		$Url,

		# SharePoint credentials
		[Microsoft.SharePoint.Client.SharePointOnlineCredentials]
		$Credentials,

		# Use web login (Allows for Multi-Factor Authentication)
		[Switch]
		$Web
	)
	process {
		# Use web login
		if ($Web) {
			Connect-PnPOnline -Url $Url -UseWebLogin
			return Get-PnPContext
		}
    
		# Use non web login
		$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
		$Context.Credentials = $Credentials
		try {
			$Context.ExecuteQuery()
		}
		catch {
			Write-Error "Invalid credentials : $($_.Exception.Message)"
			return $Null
		}
		return $Context
	}
}

<# SITE #>

function Get-SPSite() {
	<#
        .SYNOPSIS
            Returns a SharePoint site from a url and credentials
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Web])]
	param (
		# The URL of the site to get
		[String]
		$Url,

		# SharePoint Credentials
		[Microsoft.SharePoint.Client.SharePointOnlineCredentials]
		$Credentials,

		# Use web login (Allows for Multi-Factor Authentication)
		[Switch]
		$Web
	)
	process {
		# Use web login
		if ($Web) {
			Connect-PnPOnline -Url $Url -UseWebLogin
			$Context = Get-PnPContext
		}
		else {
			$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
			$Context.Credentials = $Credentials
		}
		
		$SiteWeb = $Context.Web            
		$Context.Load($SiteWeb)
		$Context.ExecuteQuery()
	
		return $SiteWeb
	}
}


function Get-SPSubSites() {
	<#
        .SYNOPSIS
            Returns all the subsites of a SharePoint site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Web[]])]
	param (
		# The site to get subsites from
		[Microsoft.SharePoint.Client.Web]
		$ParentSite
	)
	process {
		$SubSites = $ParentSite.Webs
		$ParentSite.Context.Load($SubSites)
		$ParentSite.Context.ExecuteQuery()
	
		$SubSites
	}
}


function Test-SPSite() {
	<#
        .SYNOPSIS
            Returns whether a SharePoint site exists at the given url
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The URL of the site to test
		[String]
		$Url,

		# SharePoint credentials
		[Microsoft.SharePoint.Client.SharePointOnlineCredentials]
		$Credentials
	)
	process {
		$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
		$Context.Credentials = $Credentials
		$SiteWeb = $Context.Web            
		$Context.Load($SiteWeb) 
		try {
			$Context.ExecuteQuery()
			return $True
		}
		catch {
			return $False
		}
	}
}


function Add-SPSite() {
	<#
        .SYNOPSIS
            Creates a SharePoint site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Web])]
	param (
		# The name to give the site
		[String]
		$Name,

		# The url of the site relative to the parent site
		[String]
		$RelativeUrl,

		# The parent site to add the new site to
		[Microsoft.SharePoint.Client.Web]
		$ParentSite,

		# (Optional) The template ID to base the site off of
		[Parameter(Mandatory = $False)]
		[String]
		$Template
	)
	process {
		$WebCI = New-Object Microsoft.SharePoint.Client.WebCreationInformation
		$WebCI.Title = $Name
		$WebCI.Url = $RelativeUrl
		if ($Template -ne $Null) {
			$WebCI.WebTemplate = $Template
		}
		$SubWeb = $ParentSite.Webs.Add($WebCI)
		$ParentSite.Context.ExecuteQuery()
		return $SubWeb
	}
}

function Remove-SPSite() {
	<#
        .SYNOPSIS
            Deletes a SharePoint site
    #>
	[CmdletBinding()]
	param (
		# The URL of the site to remove
		[String]
		$SiteUrl,

		# SharePoint credentials
		[Microsoft.SharePoint.Client.SharePointOnlineCredentials]
		$Credentials
	)
	process {
		#Get Web information and subsites
		$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
		$Context.Credentials = $Credentials
		$Web = $Context.Web
		$Context.Load($Web)
		$Context.Load($Web.Webs)
		$Context.executeQuery() 
	
		#Iterate through each subsite in the current web
		foreach ($Subweb in $Web.Webs) {
			#Call the function recursively to process all subsites underneath the current web
			Remove-SPSite -SiteUrl $Subweb.Url -Credentials $Credentials
		}
	
		#Delete subsite
		$Web.DeleteObject()
		$Context.ExecuteQuery()
	}
}

<# FILE #>


function Get-SPFile() {
	<#
        .SYNOPSIS
            Gets the SharePoint file at a specified URL
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.File])]
	param (
		# The server-relative URL of the file to get
		[String]
		$ServerRelativeUrl,

		# The SharePoint context in the site where the file is stored
		[Microsoft.SharePoint.Client.ClientContext]
		$Context
	)
	process {
		$File = $Context.Web.GetFileByServerRelativeUrl($ServerRelativeUrl)
		$Context.Load($File)
		$Context.ExecuteQuery()
	
		return $File
	}
}

function Test-SPFile() {
	<#
        .SYNOPSIS
            Returns whether a SharePoint file exists at a specified URL
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The server-relative URL of the file
		[String]
		$ServerRelativeUrl,

		# The SharePoint context in the site where the file is stored
		[Microsoft.SharePoint.Client.ClientContext]
		$Context
	)
	process {
		try {
			$F = Get-SPFile -ServerRelativeUrl $ServerRelativeUrl -Context $Context
			return $True
		}
		catch {
			return $False
		} 
	}
}


function Add-SPFile() {
	<#
        .SYNOPSIS
            Uploads a local file to SharePoint, if the file is above a certain size it will split it into chunks and copy them
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.File])]
	param (
		# The local file to upload
		[System.IO.FileInfo]
		$LocalFile,

		# The SharePoint folder to upload the file into
		[Microsoft.SharePoint.Client.Folder]
		$SPParentFolder,

		# (Optional) The name of the new file
		[Parameter(Mandatory = $False)]
		[String]
		$Name,
		
		# (Optional) Whether the column is a required column
		[Parameter(Mandatory = $False)]
		[Switch]
		$PrintOutput = $False
	)
	process {
		if (!$Name) {
			$Name = $LocalFile.Name
		}
	
		$FileChunkSizeInMB = 9
		$BlockSize = $FileChunkSizeInMB * 1024 * 1024
		$FileSize = $LocalFile.Length
		
		#Create Filestream
		$FileStream = New-Object IO.FileStream($LocalFile.FullName, [System.IO.FileMode]::Open)
		
		#If file is small enough, upload it all at once
		if ($FileSize -le $BlockSize) {
			#Set File Info
			$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
			$FileCreationInfo.Overwrite = $True
			$FileCreationInfo.ContentStream = $FileStream
			$FileCreationInfo.URL = $SPParentFolder.ServerRelativeUrl + "/" + $Name
			
			#Add and load file
			$SPFile = $SPParentFolder.Files.Add($FileCreationInfo)
			$SPParentFolder.Context.Load($SPFile)
			$SPParentFolder.Context.ExecuteQuery()
			
			if ($PrintOutput) {
				Write-Host "File upload complete"
			}
	
			return $SPFile
			# If file is not small enough, upload in slices
		}
		else {
			$UploadId = [GUID]::NewGuid() # Each sliced upload requires a unique ID.
			[Microsoft.SharePoint.Client.File] $Upload
			$BytesUploaded = $Null
			$BinaryReader = New-Object System.IO.BinaryReader($FileStream)
			$Buffer = New-Object System.Byte[]($BlockSize)
			$LastBuffer = $Null
			$Fileoffset = 0
			$TotalBytesRead = 0
			$BytesRead
			$First = $True
			$Last = $False
		
			# Read data from file system in blocks. 
			while (($BytesRead = $BinaryReader.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
				$TotalBytesRead = $TotalBytesRead + $BytesRead
		
				# You've reached the end of the file.
				if ($TotalBytesRead -eq $FileSize) {
					$Last = $True
					# Copy to a new buffer that has the correct size.
					$LastBuffer = New-Object System.Byte[]($BytesRead)
					[array]::Copy($Buffer, 0, $LastBuffer, 0, $BytesRead)
				}
		
				#If first slice?
				if ($First) {
					$ContentStream = New-Object System.IO.MemoryStream
					# Add an empty file.
					$FileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
					$FileInfo.ContentStream = $ContentStream
					$FileInfo.Url = $LocalFile.Name
					$FileInfo.Overwrite = $True
					$Upload = $SPParentFolder.Files.Add($FileInfo)
					$SPParentFolder.Context.Load($Upload)
		
					# Start upload by uploading the first slice.
					$Slice = [System.IO.MemoryStream]::new($Buffer) 
		
					# Call the start upload method on the first slice.
					$BytesUploaded = $Upload.StartUpload($UploadId, $Slice)
					$SPParentFolder.Context.ExecuteQuery()
		
					# fileoffset is the pointer where the next slice will be added.
					$Fileoffset = $BytesUploaded.Value
		
					# You can only start the upload once.
					$First = $False

					Write-Host "Large file upload started"
					# If not the first slice
				}
				else {
					# Get a reference to your file.
					$Upload = $SPParentFolder.Context.Web.GetFileByServerRelativeUrl($SPParentFolder.ServerRelativeUrl + [System.IO.Path]::AltDirectorySeparatorChar + $LocalFile.Name);
					# If last slice
					If ($Last) {
						$Slice = [System.IO.MemoryStream]::new($LastBuffer)
		
						# End sliced upload by calling FinishUpload.
						$Upload = $Upload.FinishUpload($UploadId, $Fileoffset, $Slice)
						$SPParentFolder.Context.Load($Upload)
						$SPParentFolder.Context.ExecuteQuery()
		
						Write-Host "Large file upload complete"
						# Return the file object for the uploaded file.
						return $Upload
						#If a middle slice
					}
					else {
						$Slice = [System.IO.MemoryStream]::new($Buffer)
		
						# Continue sliced upload.
						$BytesUploaded = $Upload.ContinueUpload($UploadId, $Fileoffset, $Slice)
						$SPParentFolder.Context.ExecuteQuery()
							
						# Update fileoffset for the next slice.
						$Fileoffset = $BytesUploaded.Value

						#Calculate percentage
						$Percentage = [math]::Round(($BytesUploaded.Value / $FileSize) * 100, 2)
						Write-Host $Percentage"%"
					}
				}
			}
		}
	}
}

function Remove-SPFile() {
	<#
        .SYNOPSIS
            Removes a SharePoint file
    #>
	[CmdletBinding()]
	param (
		# The SharePoint file to delete
		[Microsoft.SharePoint.Client.File]
		$SPFile
	)
	process {
		$SPFile.DeleteObject()
		$SPFile.Context.ExecuteQuery()
	}
}

function Copy-SPFile() {
	<#
        .SYNOPSIS
            Copies a SharePoint file externally, if the file is above a certain size it will split it into chunks and copy them
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.File])]
	param (
		# The SharePoint file to transfer
		[Microsoft.SharePoint.Client.File]
		$SPFile,

		# The destination SharePoint folder
		[Microsoft.SharePoint.Client.Folder]
		$DestinationParentFolder
	)
	process {
		$FileChunkSizeInMB = 9
		$UploadId = [GUID]::NewGuid() # Each sliced upload requires a unique ID.
		[Microsoft.SharePoint.Client.File] $Upload
		$BlockSize = $FileChunkSizeInMB * 1024 * 1024
		$FileSize = $SPFile.Length
		$DestinationParentFolder.Context.RequestTimeout = [System.Threading.Timeout]::Infinite
	
		#Handle special characters
		$Name = $SPFile.Name.Replace("#", "_")
		$Name = $Name.Replace("&", "_")
	
		if ($FileSize -le $BlockSize) {
			#Copy File Normally
			$DestinationUrl = $DestinationParentFolder.ServerRelativeUrl + "/" + $Name
			$FileStream = $SPFile.OpenBinaryStream()
			$SPFile.Context.ExecuteQuery()
			[Microsoft.SharePoint.Client.File]::SaveBinaryDirect($DestinationParentFolder.Context, $DestinationUrl, $FileStream.Value, $True)
			Write-Host "File transfer complete"
			return Get-SPFile -ServerRelativeUrl $DestinationUrl -Context $DestinationParentFolder.Context
		}
		else {
			#Copy File in Chunks
			$BytesUploaded = $Null
			$FileStream = $Null
			$FileStream = $SPFile.OpenBinaryStream()
			$SPFile.Context.ExecuteQuery()
			$BinaryReader = New-Object System.IO.BinaryReader($FileStream.Value)
			$Buffer = New-Object System.Byte[]($BlockSize)
			$LastBuffer = $Null
			$Fileoffset = 0
			$TotalBytesRead = 0
			$BytesRead
			$First = $True
			$Last = $False
	
			# Read data from file system in blocks. 
			while (($BytesRead = $BinaryReader.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
				$TotalBytesRead = $TotalBytesRead + $BytesRead
	
				# You've reached the end of the file.
				if ($TotalBytesRead -eq $FileSize) {
					$Last = $True
					# Copy to a new buffer that has the correct size.
					$LastBuffer = New-Object System.Byte[]($BytesRead)
					[array]::Copy($Buffer, 0, $LastBuffer, 0, $BytesRead)
				}
	
				if ($First) {
					$ContentStream = New-Object System.IO.MemoryStream
					# Add an empty file.
					$FileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
					$FileInfo.ContentStream = $ContentStream
					$FileInfo.Url = $Name
					$FileInfo.Overwrite = $True
					$Upload = $DestinationParentFolder.Files.Add($FileInfo)
					$DestinationParentFolder.Context.Load($Upload)
	
					# Start upload by uploading the first slice.
					$Slice = [System.IO.MemoryStream]::new($Buffer) 
	
					# Call the start upload method on the first slice.
					$BytesUploaded = $Upload.StartUpload($UploadId, $Slice)
					$DestinationParentFolder.Context.ExecuteQuery()
	
					# fileoffset is the pointer where the next slice will be added.
					$Fileoffset = $BytesUploaded.Value
	
					# You can only start the upload once.
					$First = $False

					Write-Host "Large file transfer started"
				}
				else {
					# Get a reference to your file.
					$Upload = $DestinationParentFolder.Context.Web.GetFileByServerRelativeUrl($DestinationParentFolder.ServerRelativeUrl + [System.IO.Path]::AltDirectorySeparatorChar + $SPFile.Name);
					If ($Last) {
						# Is this the last slice of data?
						$Slice = [System.IO.MemoryStream]::new($LastBuffer)
	
						# End sliced upload by calling FinishUpload.
						$Upload = $Upload.FinishUpload($UploadId, $Fileoffset, $Slice)
						$DestinationParentFolder.Context.Load($Upload)
						$DestinationParentFolder.Context.ExecuteQuery()

						Write-Host "Large file transfer complete"
	
						# Return the file object for the uploaded file.
						return $Upload
					}
					else {
						$Slice = [System.IO.MemoryStream]::new($Buffer)
	
						# Continue sliced upload.
						$BytesUploaded = $Upload.ContinueUpload($UploadId, $Fileoffset, $Slice)
						$DestinationParentFolder.Context.ExecuteQuery()
						
						# Update fileoffset for the next slice.
						$Fileoffset = $BytesUploaded.Value

						$Percentage = [math]::Round(($BytesUploaded.Value / $FileSize) * 100, 2)
						Write-Host $Percentage"%"
					}
				}
			}
		}
	}
}

<# FOLDER #>

function Get-SPFolder() {
	<#
        .SYNOPSIS
            Gets a SharePoint folder from a specified URL
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Folder])]
	param (
		# The server-relative URL of the folder to retrieve
		[String]
		$ServerRelativeUrl,

		# The SharePoint context where the folder is located
		[Microsoft.SharePoint.Client.ClientContext]
		$Context
	)
	process {
		$SPFolder = $Context.Web.GetFolderByServerRelativeUrl($ServerRelativeUrl)
		$Context.Load($SPFolder)
		$Context.ExecuteQuery()
	
		return $SPFolder
	}
}

function Test-SPFolder() {
	<#
        .SYNOPSIS
            Tests whether a SharePoint folder exists at a specified URL
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The server-relative URL of the folder to test
		[String]
		$ServerRelativeUrl,

		# The SharePoint context where the folder is located
		[Microsoft.SharePoint.Client.ClientContext]
		$Context
	)
	process {
		try {
			$Folder = Get-SPFolder -ServerRelativeUrl $ServerRelativeUrl -Context $Context
			return $True
		}
		catch {
			return $False
		}
	}
}


function Add-SPNewFolder() {
	<#
        .SYNOPSIS
            Adds a new sub folder to an existing folder
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Folder])]
	param (
		# The name of the folder to create
		[String]
		$FolderName,

		# The folder to place the new folder in
		[Microsoft.SharePoint.Client.Folder]
		$SPParentFolder
	)
	process {
		#Check if folder already exists
		$FolderURL = $SPParentFolder.ServerRelativeUrl + "\" + $FolderName
		if (Test-SPFolder -ServerRelativeUrl $FolderURL -Context $SPParentFolder.Context) {
			return $SPParentFolder.Context.Web.GetFolderByServerRelativeUrl($FolderURL)
		}
	
		#Create folder
		$SPFolder = $SPParentFolder.Folders.Add($FolderName)
		$SPParentFolder.Context.Load($SPFolder)
		$SPParentFolder.Context.ExecuteQuery()
		$SPFolder.Update()
		$SPParentFolder.Context.ExecuteQuery()
	
		return $SPFolder
	}
}


function Add-SPFolderFromLocalFolder() {
	<#
        .SYNOPSIS
            Uploads a local folder and all its contents into a SharePoint folder
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Folder])]
	param (
		[System.IO.DirectoryInfo]
		$LocalFolder,

		[Microsoft.SharePoint.Client.Folder]
		$SPParentFolder
	)
	process {
		#Get/Create folder
		$SPfolderURL = $SPParentFolder.ServerRelativeUrl + "/" + $LocalFolder.Name
		try { 
			$SPFolder = Get-SPFolder -ServerRelativeUrl $SPfolderURL -Context $SPParentFolder.Context 
		}
		catch { 
			$SPFolder = Add-SPNewFolder -FolderName $LocalFolder.Name -SPParentFolder $SPParentFolder
		}
	
		#Get contents of local folder
		$Contents = Get-ChildItem -Path $LocalFolder.FullName
		foreach ($Item in $Contents) {
			#If item is a folder, create the folder and recursively copy the contents of that folder
			if ($Item.PSIsContainer) {
				$NewLocalFolder = Get-Item ($LocalFolder.FullName + "\" + $Item.Name)
				$F = Add-SPFolderFromLocalFolder -LocalFolder $NewLocalFolder -SPParentFolder $SPFolder
			}
			else {
				$F = Add-SPFile -LocalFile $Item -SPParentFolder $SPFolder
			}
		}
		
		$SPParentFolder.Context.Load($SPFolder)
		$SPParentFolder.Context.ExecuteQuery()
		
		return $SPFolder
	}
}


function Remove-SPFolder() {
	<#
        .SYNOPSIS
            Deletes a SharePoint folder
    #>
	[CmdletBinding()]
	param (
		# The folder to delete
		[Microsoft.SharePoint.Client.Folder]
		$SPFolder
	)
	process {
		$SPFolder.DeleteObject()
		$SPFolder.Context.ExecuteQuery()
	}
}

function Copy-SPFolder() {
	<#
        .SYNOPSIS
            Copies a SharePoint folder externally
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Folder])]
	param (
		# The SharePoint folder to copy
		[Microsoft.SharePoint.Client.Folder]
		$SourceFolder,

		# The SharePoint folder to copy into
		[Microsoft.SharePoint.Client.Folder]
		$DestinationParentFolder
	)
	process {
		#Load source folder
		$SourceFolder.Context.Load($SourceFolder.Files)
		$SourceFolder.Context.Load($SourceFolder.Folders)
		$SourceFolder.Context.ExecuteQuery()
	
		#Create destination folder
		$DestinationFolder = Add-SPNewFolder -FolderName $SourceFolder.Name -SPParentFolder $DestinationParentFolder
	
		foreach ($File in $SourceFolder.Files) {
			#Check if List Item Exists
			$SourceFolder.Context.Load($File.ListItemAllFields)
			$SourceFolder.Context.ExecuteQuery()
			if ($Null -eq $File.ListItemAllFields.Id) { continue }
			
			#Create unused variable to prevent error
			$F = Copy-SPFile -spFile $File -destinationParentFolder $DestinationFolder
		}
	
		foreach ($Folder in $SourceFolder.Folders) {
			#Check if List Item Exists
			$SourceFolder.Context.Load($Folder.ListItemAllFields)
			$SourceFolder.Context.ExecuteQuery()
			if ($Null -eq $Folder.ListItemAllFields.Id) { continue }
	
			#Create unused variable to prevent error
			$F = Copy-SPFolder -sourceFolder $Folder -destinationParentFolder $DestinationFolder
		}
	
		return $DestinationFolder
	}
}

function Get-SPFilesInFolder() {
	<#
        .SYNOPSIS
            Gets a SharePoint folder from a specified URL
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.File[]])]
	param (
		# The folder to retrieve files from
		[Microsoft.SharePoint.Client.Folder]
		$SPFolder,

		# Whether to include files in sub folders
		[Parameter(Mandatory = $False)]
		[Switch]
		$Recurse = $False
	)
	process {
		$Files = @()

		#Load folder
		$SPFolder.Context.Load($SPFolder.Files)
		$SPFolder.Context.Load($SPFolder.Folders)
		$SPFolder.Context.ExecuteQuery()
	
		foreach ($File in $SPFolder.Files) {
			# Add to list
			$Files += $File
		}
	
		if ($Recurse) {
			foreach ($SubFolder in $SPFolder.Folders) {
				#Create unused variable to prevent error
				$Files += (Get-SPFilesInFolder -SPFolder $SubFolder -Recurse)
			}
		}
	
		return $Files
	}
}

<# LIST #>

function Get-SPList() {
	<#
        .SYNOPSIS
            Gets a SharePoint list by name from a specified SharePoint site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.List])]
	param (
		# The name of the SharePoint list to get
		[String]
		$ListName,

		# The site where the list exists
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$List = $Site.Lists.GetByTitle($ListName)
		$Site.Context.Load($List)
		$Site.Context.ExecuteQuery()
		
		return $List
	}
}


function Test-SPList() {
	<#
        .SYNOPSIS
            Tests whether a SharePoint list within a site exists with a given name
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The name of the list to check for
		[String]
		$ListName,

		# The site to look for the list in
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		try {
			$L = Get-SPList -ListName $ListName -Site $Site
			return $True
		}
		catch {
			return $False
		}
	}
}


function Add-SPList() {
	<#
        .SYNOPSIS
            Adds a new SharePoint list to a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.List])]
	param (
		# The name of the list
		[String]
		$ListName,

		# The SharePoint site to add the list to
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
		$ListInfo.Title = $ListName
		$ListInfo.TemplateType = "100"
		$List = $Site.Lists.Add($ListInfo)
		$List.Description = $ListName
		$List.Update()
		$Site.Context.ExecuteQuery()
	
		return $List
	}
}


function Remove-SPList() {
	<#
        .SYNOPSIS
            Removes a SharePoint list
    #>
	[CmdletBinding()]
	param (
		# The list to remove
		[Microsoft.SharePoint.Client.List]
		$List
	)
	process {
		$List.DeleteObject()
		$List.Context.ExecuteQuery()
	}
}


function Get-SPListsInSite() {
	<#
        .SYNOPSIS
            Gets all SharePoint lists in a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.List[]])]
	param (
		# The SharePoint site to get the lists from
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$Lists = $Site.Lists
		$Site.Context.Load($Lists)
		$Site.Context.ExecuteQuery()
	
		return $Lists
	}
}


function Get-SPLibrariesInSite() {
	<#
        .SYNOPSIS
            Gets all SharePoint libraries in a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.List[]])]
	param (
		# The SharePoint site to get the libraries from
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$Libraries = @()
	
		$Lists = $Site.Lists
		$Site.Context.Load($Lists)
		$Site.Context.ExecuteQuery()
	
		#Check if list is a library
		foreach ($List in $Lists) {
			if ($List.BaseTemplate -eq 101) {
				$Libraries += $List
			}
		}
	
		return $Libraries
	}
}

function Add-SPLibrary() {
	<#
        .SYNOPSIS
            Adds a new SharePoint list to a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.List])]
	param (
		# The name of the list
		[String]
		$LibraryName,

		# The SharePoint site to add the list to
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
		$ListInfo.Title = $LibraryName
		$ListInfo.TemplateType = "101"
		$List = $Site.Lists.Add($ListInfo)
		$List.Description = $LibraryName
		$List.Update()
		$Site.Context.ExecuteQuery()
	
		return $List
	}
}

<# LIST ITEM #>


function Get-SPListItems() {
	<#
        .SYNOPSIS
            Gets list items in a SharePoint list
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ListItem[]])]
	param (
		# The SharePoint list to get the list items from
		[Microsoft.SharePoint.Client.List]
		$List,
		
		# (Optional) The query to specify which list items to retrieve
		[Parameter(Mandatory = $False)]
		[String]
		$Query
	)
	process {
		$CamlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
		if ($Query) {
			$CamlQuery.ViewXML = $Query
		}
		$ListItems = $List.GetItems($CamlQuery)
		$List.Context.Load($ListItems)
		$List.Context.ExecuteQuery()
	
		return $ListItems
	}
}


function Get-SPListItemsByField() {
	<#
        .SYNOPSIS
            Gets list items in a SharePoint list with a specific property
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ListItem[]])]
	param (
		# The SharePoint list to get the list items from
		[Microsoft.SharePoint.Client.List]
		$List,

		# The field to check
		[String]
		$Field,

		# The value of the field to check against
		[String]
		$Value
	)
	process {
		$Query = New-Object Microsoft.SharePoint.Client.CamlQuery 
		$Query.ViewXML = "<View><Query><Where><Eq><FieldRef Name='$Field' /><Value Type='Text'>$Value</Value></Eq></Where></Query></View>"
		$ListItems = $List.GetItems($Query)
		$List.Context.Load($ListItems)
		$List.Context.ExecuteQuery()
	
		return $ListItems
	}
}


function Test-SPListItemByField() {
	<#
        .SYNOPSIS
            Tests whether list items with a specific property exist in a SharePoint list
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The SharePoint list to get the list items from
		[Microsoft.SharePoint.Client.List]
		$List,
		
		# The field to check
		[String]
		$Field,
		
		# The value of the field to check against
		[String]
		$Value
	)
	process {
		$Query = New-Object Microsoft.SharePoint.Client.CamlQuery 
		$Query.ViewXML = "<View><Query><Where><Eq><FieldRef Name='$Field' /><Value Type='Text'>$Value</Value></Eq></Where></Query></View>"
		$ListItems = $List.GetItems($Query)
		$List.Context.Load($ListItems)
		$List.Context.ExecuteQuery()
	
		return $ListItems.Count -gt 0
	}
}


function Add-SPListItem() {
	<#
        .SYNOPSIS
            Adds a new list item to a SharePoint list
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ListItem])]
	param (
		# The list to add the new list item to
		[Microsoft.SharePoint.Client.List]
		$List
	)
	process {
		$ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
		$NewListItem = $List.AddItem($ListItemCreationInformation)
		$List.Context.ExecuteQuery()
	
		return $NewListItem
	}
}


function Update-SPListItem() {
	<#
        .SYNOPSIS
            Updates a list item
    #>
	[CmdletBinding()]
	param (
		# The list item to update
		[Microsoft.SharePoint.Client.ListItem]
		$ListItem
	)
	process {
		$ListItem.Update()
		$ListItem.Context.Load($ListItem)
		$ListItem.Context.ExecuteQuery()
	}
}


function Remove-SPListItem() {
	<#
        .SYNOPSIS
            Deletes a list item
    #>
	[CmdletBinding()]
	param (
		# The list item to delete
		[Microsoft.SharePoint.Client.ListItem]
		$ListItem
	)
	process {
		$ListItem.DeleteObject()
		$ListItem.Context.ExecuteQuery()
	}
}

<# TERMS #>


function Get-SPTermSetByName() {
	<#
        .SYNOPSIS
            Gets a SharePoint term set with a specified name within a specified group
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Taxonomy.TermSet])]
	param (
		# The name of the group that the term set is within
		[String]
		$GroupName,

		# The name of the term set to get
		[String]
		$TermSetName,

		# The context where the term set is held
		[Microsoft.SharePoint.Client.ClientContext]
		$Context
	)
	process {
		#Bind to MMS
		$MMS = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($Context)
		$Context.Load($MMS)
		$Context.ExecuteQuery()
		#Bind to Term Stores
		$TermStores = $MMS.TermStores
		$Context.Load($TermStores)
		$Context.ExecuteQuery()
		#Bind to Term Store
		$TermStore = $TermStores[0]
		$Context.Load($TermStore)
		$Context.ExecuteQuery()
		#Bind to Group
		$Group = $TermStore.Groups.GetByName($GroupName)
		$Context.Load($Group)
		$Context.ExecuteQuery()
		#Bind to Term Set
		$TermSet = $Group.TermSets.GetByName($TermSetName)
		$Context.Load($TermSet)
		$Context.ExecuteQuery()
	
		return $TermSet
	}
}


function Get-SPTermByName() {
	<#
        .SYNOPSIS
            Gets a term from a SharePoint term set by name
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Taxonomy.Term])]
	param (
		# The term set where the term is contained
		[Microsoft.SharePoint.Client.Taxonomy.TermSet]
		$TermSet,

		# The name of the term
		[String]
		$TermName
	)
	process {
		$Terms = $TermSet.GetAllTerms()
		$TermSet.Context.Load($Terms)
		$TermSet.Context.ExecuteQuery()
	
		#Find Term
		foreach ($Term in $Terms) {
			if ($Term.Name -eq $TermName) {
				return $Term
			}
		}
	
		return $Null
	}
}


function Test-SPTermByName() {
	<#
        .SYNOPSIS
            Tests whether a term with a specified name exists within a SharePoint term set
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The term set to look for the term in
		[Microsoft.SharePoint.Client.Taxonomy.TermSet]
		$TermSet,

		# The name of the term to look for
		[String]
		$TermName
	)
	process {
		if ($Null -eq (Get-SPTermByName -TermSet $TermSet -TermName $TermName)) {
			return $False
		}
		else {
			return $True
		}
	}
}


function Add-SPTerm() {
	<#
        .SYNOPSIS
            Adds a new term to a term set
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Taxonomy.Term])]
	param (
		# The term set to add the term to
		[Microsoft.SharePoint.Client.Taxonomy.TermSet]
		$TermSet,

		# The name of the term to add
		[String]
		$TermName
	)
	process {
		$NewTerm = $TermSet.CreateTerm($TermName, 1033, [System.Guid]::NewGuid().toString())
		$TermSet.Context.Load($NewTerm)
		$TermSet.Context.ExecuteQuery()
	
		return $NewTerm
	}
}


function Remove-SPTerm() {
	<#
        .SYNOPSIS
            Deletes a term
    #>
	[CmdletBinding()]
	param (
		# The term to delete
		[Microsoft.SharePoint.Client.Taxonomy.Term]
		$Term
	)
	process {
		$Term.DeleteObject()
		$Term.Context.ExecuteQuery()
	}
}

<# COLUMNS #>


function Get-SPColumn() {
	<#
        .SYNOPSIS
            Gets a site column by name
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Field])]
	param (
		#The name of the column to get
		[String]
		$ColumnName,

		# The site to get the column from
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$Fields = $Site.Fields
		$Site.Context.Load($Fields)
		$Site.Context.ExecuteQuery()
		$Column = $Fields.GetByInternalNameOrTitle($ColumnName)
	
		return $Column
	}
}


function Test-SPColumn() {
	<#
        .SYNOPSIS
            Tests whether a site column exists by a given name
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The name of the column to look for
		[String]
		$ColumnName,

		# The site to look for the column in
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		try {
			$C = Get-SPColumn -ColumnName $ColumnName -Site $Site
			return $True
		}
		catch {
			return $False
		}
	}
}


function Add-SPColumn() {
	<#
        .SYNOPSIS
            Adds a new column to a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Field])]
	param (
		# The name of the column
		[String]
		$ColumnName,

		# The data type of the column
		# Boolean / DateTime / Number / Text / User
		# For more check here: https://docs.microsoft.com/en-us/sharepoint/dev/schema/field-element-field
		[String]
		$ColumnType,

		# The site to add the column to
		[Microsoft.SharePoint.Client.Web]
		$Site,
		
		# (Optional) The group to add the column to
		[Parameter(Mandatory = $False)]
		[String]
		$ColumnGroup = "Custom Columns",

		# (Optional) Whether the column is a required column
		[Parameter(Mandatory = $False)]
		[Switch]
		$Required = $False
	)
	process {
		$Fields = $Site.Fields
		$Site.Context.Load($Fields)
		$Site.Context.ExecuteQuery()
	
		$FieldXML = "<Field Type='$ColumnType' DisplayName='$ColumnName' Name='$ColumnName' required='$Required' Group='$ColumnGroup'/>"
		$Column = $Site.Context.Web.Fields.AddFieldAsXml($FieldXML, $True, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
		$Site.Context.ExecuteQuery()
	
		return $Column
	}
}


function Remove-SPColumn() {
	<#
        .SYNOPSIS
            Deletes a site column
    #>
	[CmdletBinding()]
	param (
		# The column to delete
		[Microsoft.SharePoint.Client.Field]
		$Column
	)
	process {
		$Column.DeleteObject()
		$Column.Context.ExecuteQuery()
	}
}

<# CONTENT TYPES #>


function Get-SPContentTypesInSite() {
	<#
        .SYNOPSIS
            Gets all content types in a SharePoint site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ContentType[]])]
	param (
		# The site to get the content types from
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$ContentTypes = $Site.ContentTypes
		$Site.Context.Load($ContentTypes)
		$Site.Context.ExecuteQuery()
	
		return $ContentTypes
	}
}


function Get-SPSiteContentTypeByName() {
	<#
        .SYNOPSIS
            Gets a site content type by name
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ContentType])]
	param (
		# The name of the content type
		[String]
		$ContentTypeName,

		# The site to get the content type from
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$ContentTypes = $Site.ContentTypes
		$Site.Context.Load($ContentTypes)
		$Site.Context.ExecuteQuery()
	
		#Find Content Type
		foreach ($ContentType in $ContentTypes) {
			if ($ContentType.Name -eq $ContentTypeName) {
				return $ContentType
			}
		}
	
		return $Null
	}
}


function Test-SPSiteContentTypeByName() {
	<#
        .SYNOPSIS
            Tests whether a content type with a given name exists within a site
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The name of the content type
		[String]
		$ContentTypeName,

		# The site to look for the content type in
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$ContentType = Get-SPSiteContentTypeByName -ContentTypeName $ContentTypeName -Site $Site
		if ($Null -eq $ContentType) {
			return $False
		}
		else {
			return $True
		}
	}
}


function Add-SPContentTypeToSite() {
	<#
        .SYNOPSIS
            Adds a new content type to a site
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ContentType])]
	param (
		# The name of the new content type
		[String]
		$ContentTypeName,

		# The columns to add to the content type by name
		[String[]]
		$Columns,

		# The content type to inherit from
		[Microsoft.SharePoint.Client.ContentType]
		$ParentContentType,

		# (Optional) The group to place the content type in
		[Parameter(Mandatory = $False)]
		[String]
		$ContentTypeGroup = "Custom Content Types"
	)
	process {
		#Get content type collection from site
		$ContentTypes = $ParentContentType.Context.Web.ContentTypes
		$ParentContentType.Context.Load($ContentTypes)
		$ParentContentType.Context.ExecuteQuery()
	
		#Get field collection from site
		$Fields = $ParentContentType.Context.Web.Fields
		$ParentContentType.Context.Load($Fields)
		$ParentContentType.Context.ExecuteQuery()
		
		#Set new content type information
		$CTCreationInformation = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation
		$CTCreationInformation.Name = $ContentTypeName
		$CTCreationInformation.ParentContentType = $ParentContentType
		$CTCreationInformation.Group = $ContentTypeGroup
	
		#Create content type
		$ContentType = $ContentTypes.Add($CTCreationInformation)
		$ParentContentType.Context.Load($ContentType)
		$ParentContentType.Context.ExecuteQuery()
	
		#Add columns from site fields
		foreach ($Column in $Columns) {
			$Field = $Fields.GetByInternalNameOrTitle($Column)
			$FieldLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
			$FieldLink.Field = $Field
			$C = $ContentType.FieldLinks.Add($FieldLink)
		}
		$ContentType.Update($True)
		$ParentContentType.Context.ExecuteQuery()
	
		return $ContentType
	}
}


function Remove-SPContentType() {
	<#
        .SYNOPSIS
            Deletes a content type
    #>
	[CmdletBinding()]
	param (
		# The content type to delete
		[Microsoft.SharePoint.Client.ContentType]
		$ContentType
	)
	process {
		$ContentType.DeleteObject()
		$ContentType.Context.ExecuteQuery()
	}
}

function Add-SPContentTypeToList() {
	<#
        .SYNOPSIS
            Adds a content type to a SharePoint list
    #>
	[CmdletBinding()]
	param (
		# The content type to add by name
		[String]
		$ContentTypeName,

		# The list to add the content type to
		[Microsoft.SharePoint.Client.List]
		$List,

		# Whether the content type should be read only
		[Parameter(Mandatory = $False)]
		[Switch]
		$SetReadable
	)
	process {
		$ContentType = Get-SPSiteContentTypeByName -ContentTypeName $ContentTypeName -Context $List.Context
		if ($SetReadable) { 
			$ContentType.ReadOnly = $False
		}
		$List.ContentTypesEnabled = $True
		$C = $List.ContentTypes.AddExistingContentType($ContentType)
		$List.Update()
		$List.Context.ExecuteQuery()
	}
}

function Get-SPContentTypesInList() {
	<#
        .SYNOPSIS
            Returns all the content types attached to a list
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ContentType[]])]
	param (
		# The list to get the content types from
		[Microsoft.SharePoint.Client.List]
		$List
	)
	process {
		$ContentTypes = $List.ContentTypes
		$List.Context.Load($ContentTypes)
		$List.Context.ExecuteQuery()
	
		return $ContentTypes
	}
}


function Get-SPListContentTypeByName() {
	<#
        .SYNOPSIS
            Gets a content type from a list by name
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.ContentType])]
	param (
		# The name of the content type
		[String]
		$ContentTypeName,

		# The list to get the content type from
		[Microsoft.SharePoint.Client.List]
		$List
	) 
	process {
		$ContentTypes = $List.ContentTypes
		$List.Context.Load($ContentTypes)
		$List.Context.ExecuteQuery()
	
		foreach ($ContentType in $ContentTypes) {
			if ($ContentType.Name -eq $ContentTypeName) {
				return $ContentType
			}
		}
	
		return $Null
	}
}


function Test-SPListContentTypeByName() {
	<#
        .SYNOPSIS
            Tests whether a content type with a given name exists within a list
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The name of the content type
		[String]
		$ContentTypeName,

		# The list to look for the content type in
		[Microsoft.SharePoint.Client.List]
		$List
	)
	process {
		$ContentType = Get-SPListContentTypeByName -ContentTypeName $ContentTypeName -List $List
		if ($Null -eq $ContentType) {
			return $False
		}
		else {
			return $True
		}
	}
}


function Remove-SPContentTypeFromList() {
	<#
        .SYNOPSIS
            Removes a content type from a list
    #>
	[CmdletBinding()]
	param (
		# The name of the content type to remove
		[String]
		$ContentTypeName,

		# The list to remove the content type from
		[Microsoft.SharePoint.Client.List]
		$List
	) 
	process {
		$ContentTypes = $List.ContentTypes
		$List.Context.Load($ContentTypes)
		$List.Context.ExecuteQuery()
	
		foreach ($ContentType in $ContentTypes) {
			if ($ContentType.Name -eq $ContentTypeName) {
				$ContentType.DeleteObject()
				break
			}
		}
	
		$List.Context.ExecuteQuery()
	}
}

function Get-SPListItemContentType() {
	<#
        .SYNOPSIS
            Gets the content type of a list item
    #>
	[CmdletBinding()]
	param (
		# The list to remove the content type from
		[Microsoft.SharePoint.Client.ListItem]
		$ListItem
	) 
	process {
		$List = $ListItem.ParentList
		$ListItem.Context.Load($List)
		$ListItem.Context.ExecuteQuery()

		$ContentTypes = Get-SPContentTypesInList -List $List

		foreach ($ContentType in $ContentTypes) {        
			if ($ContentType.ID.ToString() -eq $ListItem["ContentTypeId"].ToString()) {
				return $ContentType
			}
		}

		return $Null
	}
}

<# PERMISSIONS #>

function Get-SPSiteGroupByName() {
	<#
        .SYNOPSIS
            Returns a SharePoint group from a site by name
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Web])]
	param (
		# The site to get the group from
		[Microsoft.SharePoint.Client.Web]
		$Site,

		# The name of the group to get
		[String]
		$GroupName
	)
	process {
		$SiteGroups = $Site.SiteGroups
		$Site.Context.Load($siteGroups)
		$Site.Context.ExecuteQuery()
		
		return $SiteGroups | where { $_.Title -eq $user.Title }
	}
}

function Get-SPRoleUser() {
	<#
        .SYNOPSIS
            Gets a user from a role
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Principal])]
	param (
		# The role assignment to extract the user from
		[Microsoft.SharePoint.Client.RoleAssignment]
		$RoleAssignment
	)
	process {
		$RoleAssignment.Context.Load($RoleAssignment.Member)
		$RoleAssignment.Context.ExecuteQuery()
	
		return $RoleAssignment.Member
	}
}


function Get-SPRolePermissions() {
	<#
        .SYNOPSIS
            Gets the permission level of a role
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.RoleDefinition])]
	param (
		# The role assignment to get the permission level of
		[Microsoft.SharePoint.Client.RoleAssignment]
		$RoleAssignment
	)
	process {
		$RoleAssignment.Context.Load($RoleAssignment.RoleDefinitionBindings)
		$RoleAssignment.Context.ExecuteQuery()
	
		return $RoleAssignment.RoleDefinitionBindings
	}
}

function Get-SPSitePermissions() {
	<#
        .SYNOPSIS
            Gets the permissions of an object
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.RoleAssignment[]])]
	param (
		# The object to get the permissions of
		[Microsoft.SharePoint.Client.Web]
		$Site
	)
	process {
		$RoleAssignments = $Site.RoleAssignments
		$Site.Context.Load($RoleAssignments)
		$Site.Context.ExecuteQuery()
	
		return $RoleAssignments
	}
}

function Get-SPObjectPermissions() {
	<#
        .SYNOPSIS
            Gets the permissions of an object
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.RoleAssignment[]])]
	param (
		# The object to get the permissions of
		[Microsoft.SharePoint.Client.ClientObject]
		$Object
	)
	process {
		$RoleAssignments = $Object.RoleAssignments
		$Object.Context.Load($RoleAssignments)
		$Object.Context.ExecuteQuery()
	
		return $RoleAssignments
	}
}


function Set-SPObjectPermission() {
	<#
        .SYNOPSIS
            Adds a new permission to an object
    #>
	[CmdletBinding()]
	param (
		# The object to add the permission to
		[Microsoft.SharePoint.Client.ClientObject]
		$Object,

		# The permission level to give
		[String]
		$RoleName,

		# The user to give the permission to
		[String]
		$UserLoginName
	) 
	process {
		$Site = $Object.Context.Web
		$Object.Context.Load($Site)
		$Object.Context.ExecuteQuery()

		# Get the permission level and user by name
		$NewRole = $Site.RoleDefinitions.GetByName($RoleName)
		$NewMember = $Site.EnsureUser($UserLoginName)  
		$Object.Context.ExecuteQuery()
		
		# Break the inheritance if it exists
		$Object.BreakRoleInheritance($False, $False)  
		$Object.Context.Load($Object)
		
		#Create the new role assignment
		$NewRoleAssignment = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Object.Context)
		$NewRoleAssignment.Add($NewRole)
		$NewPermission = $Object.RoleAssignments.Add($NewMember, $NewRoleAssignment) 
		$Object.Context.Load($NewPermission)
		$Object.Context.ExecuteQuery()
	}
}

function Remove-SPObjectPermission() {
	<#
        .SYNOPSIS
            Removes an object permission by user
    #>
	[CmdletBinding()]
	param (
		# The site to remove the permission from
		[Microsoft.SharePoint.Client.ClientObject]
		$Object,

		# The user to remove
		[String]
		$UserLoginName
	)
	process {
		$Object.BreakRoleInheritance($False, $False)  
		$Object.Context.Load($Object)
	
		$Permission = $Object.RoleAssignments.GetByPrincipal($Object.Context.Web.EnsureUser($UserLoginName))
		$Object.Context.Load($Permission)
		try {
			$Object.Context.ExecuteQuery()
		}
		catch {
			Write-Host "User" $UserLoginName "does not exist." -ForegroundColor Red
		}
		
		$Permission.DeleteObject()
	
		$Object.Context.ExecuteQuery()
	}
}


function Clear-SPObjectPermissions() {
	<#
        .SYNOPSIS
            Removes all permission from an object
    #>
	[CmdletBinding()]
	param (
		# The site to clear the permissions of
		[Microsoft.SharePoint.Client.ClientObject]
		$Object
	)
	process {
		$Object.BreakRoleInheritance($False, $False)  
		$Object.Context.Load($Object)
		
		$Object.Context.ExecuteQuery()
		
		$Count = $Object.RoleAssignments.Count - 1;
		for ($i = $Count; $i -ge 0; $i--) {
			$Object.RoleAssignments[$i].DeleteObject()
		}
	
		$Object.Context.ExecuteQuery()
	}
}

function Test-SPObjectHasUniquePermission() {
	<#
        .SYNOPSIS
            Returns whether an object has unique permissions
    #>
	[CmdletBinding()]
	[OutputType([Boolean])]
	param (
		# The object to test
		[Microsoft.SharePoint.Client.ClientObject]
		$Object
	)
	process {
		if ($Object -is [Microsoft.SharePoint.Client.File] -or $Object -is [Microsoft.SharePoint.Client.Folder]) {
			$Object = $Object.ListItemAllFields
		}

		Invoke-LoadMethod -Object $Object -PropertyName "HasUniqueRoleAssignments"
	
		return $Object.HasUniqueRoleAssignments
	}
}


function Get-SPWorkflowsInObject() {
	<#
        .SYNOPSIS
            Returns workflows in an object
    #>
	[CmdletBinding()]
	[OutputType([Microsoft.SharePoint.Client.Workflow.WorkflowAssociation[]])]
	param (
		# The object to get the workflows of
		[Microsoft.SharePoint.Client.ClientObject]
		$Object
	)
	process {
		$Workflows = $Object.WorkflowAssociations
		$Object.Context.Load($Workflows)
		$Object.Context.ExecuteQuery()
	
		return $Workflows
	}
}

<# UTILITY #>

function Invoke-LoadMethod() {
	<#
        .SYNOPSIS
            Retrieves a property from a SharePoint object
    #>
	[CmdletBinding()]
	param (
		# The object to retrieve the property for
		[Microsoft.SharePoint.Client.ClientObject]
		$Object,

		# The name of the parameter to retrieve
		[String]
		$PropertyName
	)
	process {
		$Load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load")
		$Type = $Object.GetType()
		$ClientLoad = $Load.MakeGenericMethod($Type)
		$Parameter = [System.Linq.Expressions.Expression]::Parameter(($Type), $Type.Name)
		$Expression = [System.Linq.Expressions.Expression]::Lambda(
			[System.Linq.Expressions.Expression]::Convert(
				[System.Linq.Expressions.Expression]::PropertyOrField($Parameter, $PropertyName),
				[System.Object]
			),
			$($Parameter)
		)
		$ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
		$ExpressionArray.SetValue($Expression, 0)
		$ClientLoad.Invoke($Object.Context, @($Object, $ExpressionArray))
		$Object.Context.ExecuteQuery()
	}
}