Function SFTPUploadFiles {
	Param(
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $Username = $(throw "Username parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $Password = $(throw "Password parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $HostName = $(throw "HostName parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $RemotePath = $(throw "RemotePath parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $LocalPath = $(throw "LocalPath parameter is required"),
		[Parameter(Mandatory=$true)]
		[ValidateNotNull()]
		[string] $SshHostKeyFingerprint = $(throw "SshHostKeyFingerprint parameter is required"),
		$FileMask="",
		$Remove=$false,
		$TransferMode=""
	)
	if( -not (Test-Path $LocalPath)) {
		throw("ERROR: Unable to locate LocalPath (path=${LocalPath})")
	}


	$Invocation = (Get-Variable MyInvocation -Scope 1).Value
	$SftpModuleDirectory = Split-Path $Invocation.MyCommand.Path
	
    [Reflection.Assembly]::LoadFrom("${SftpModuleDirectory}\lib\WinSCP.dll") | Out-Null
 
    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions
    $sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
    $sessionOptions.HostName = $HostName
    $sessionOptions.UserName = $Username
    $sessionOptions.Password = $Password
    $sessionOptions.SshHostKeyFingerprint = $SshHostKeyFingerprint #"ssh-rsa 1024 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx"
 
    $session = New-Object WinSCP.Session
 
	Write-Host "--------------------------------"
	Write-Host "-- Connection Details"
	Write-Host "--------------------------------"
	Write-Host "HostName: " $HostName
	Write-Host "Username: " $Username
	Write-Host "RemotePath: " $RemotePath
	Write-Host "LocalPath: " $LocalPath
	if($FileMask -ne "") {
		Write-Host "FileMask: " $FileMask
	}
	Write-Host "--------------------------------"
	
    try
    {
		$session.ExecutablePath = "${SftpModuleDirectory}\lib\WinSCP.exe"
		
		Write-Host "- attempting to connect to ftp server"
        # connect to FTP session
		try {
	    $session.Open($sessionOptions)
		} catch {
		
			if($_.Exception.ToString().Contains("Host key wasn't verified!")) {
				throw("invalid SshHostKeyFingerprint, unable to open session to FTP (host=${HostName}, SshHostKeyFingerprint=${SshHostKeyFingerprint})")
			}		
			elseif($_.Exception.ToString().Contains("No supported authentication methods available")) {
				throw("Unable to open session to FTP (host=${HostName}, username=${Username})")
			}		
		}
	
		if (-not ($session.FileExists($remotePath))) {
			throw("The RemotePath does not exist on FTP (path=${RemotePath})")
		}
		
		# set optional params
		$TransferOptions = New-Object WinSCP.TransferOptions
		if($FileMask -ne "" -and $FileMask -ne $null) {
			$TransferOptions.FileMask = $FileMask
		}
		$TransferModeObj = New-Object WinSCP.TransferMode
		if($TransferMode -ieq "ascii") {
			$TransferOptions.TransferMode = "Ascii"
		}
		elseif($TransferMode -ieq "automatic") {
			$TransferOptions.TransferMode = "Automatic"
		}
					
		Write-Host "- begin upload of files"
		# execute File Upload
		$result = $session.PutFiles($localPath, $remotePath, $remove, $TransferOptions)


	    # Throw on any error
	    $result.Check()
		
		Write-Host "--------------------------------"
		Write-Host "-- Transfer Results"
		Write-Host "--------------------------------"
		# output reult summary
		if($result.IsSuccess) {
			Write-Host "STATUS: Success"
		} else {
			Write-Host "STATUS: Failure"
		}
		Write-Host "FILE UPLOADS: " ($result.Transfers | Measure).count
		Write-Host "FAIELD UPLOADS: " ($result.Failures | Measure).count
		Write-Host "--------------------------------"
		
		
		if(($result.Transfers | Measure).count -gt 0) {
			Write-Host "- Files successfully transfered:"
			$result.Transfers | % {
				Write-Host "-- $($_.Destination)"
			}
		}
		if(($result.Failures | Measure).count -gt 0) {
			Write-Host "-Failed file transfers:"
			$result.Failures | % {
				Write-Host "-- $($_.FileName    )"
			}
		}
		if(-not $result.IsSuccess) {
			throw("FTP transfer ended unsuccessfully")
		}
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
}
