<#
.Synopsis
   Choose an organizational unit from a GUI
.DESCRIPTION
   Launches a windows form where you can choose an organizational unit from
   a treeview. You can change the domain or add an organizational unit from 
   the context menu.
.EXAMPLE
   Choose-ADOrganizationalUnit -HideNewOUFeature
   This command will show the form containing the OU structure of the
   current users domain. The New OU feature is hidden from the context menu.
.EXAMPLE
   Choose-ADOrganizationalUnit -Domain childdomain.contoso.com -AdvancedFeatures
   This command will show the form containing the OU structure of the childdomain 
   of contoso.com. Also the distinguished name can be used (DC=CHILDDOMAIN,DC=CONTOSO,DC=COM).
   Advanced features are shown at startup.
.EXAMPLE
   Choose-ADOrganizationalUnit -Domain contoso.com -Credential CONTOSO\AdminUser
   This command will show the form containing the OU structure of the
   CONTOSO.COM domain using alternate credentials. This can also be used from a
   computer that is not domain-joined.
.EXAMPLE
   Choose-ADOrganizationalUnit -RootOU OU=Finance,OU=Departments,DC=CONTOSO,DC=COM
   This command will show the form containing the OU structure of the
   CONTOSO.COM domain using the Finance OU as root.
.OUTPUTS
   PowerShell object with Name and Distinguished Name of chosen organizational unit.
.NOTES
   Author : Michaja van der Zouwen
   version: 2.3
   Date   : 14-06-2018

   New in this version: 

   *  Change domain option in context menu disabled when only one domain is detected.
   *  Fixed an issue with domains starting with a 'D'
   *  Added a Change Farm button to the Change Domain GUI. 

   Special thanks to @danSpotter for bringing these issues to light!
.LINK
   https://itmicah.wordpress.com/2016/03/29/active-directory-ou-picker-revisited/
#>
function Choose-ADOrganizationalUnit
{
    [CmdletBinding()]
    Param
    (
        #FQDN or Distinguished Name of the domain you want to use
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Alias("DistinguishedName")]
	    [string]
        $Domain,

        #Distinghuished name of an OU you want to serve as root
        [Parameter(Mandatory=$false,
                       Position=1)]
        [string]
        $RootOU,
	
	    #Credentials for connecting to ActiveDirectory domain
	    $Credential,

        #Enable Advanced features on startup
        [switch]
        $AdvancedFeatures,
	
	    #Hide the ability to create OU's
        [switch]
        $HideNewOUFeature,

        #Add checkboxes so multiple objects can be selected
        [switch]
        $MultiSelect
    )

    
	#region Import the Assemblies

	[void][reflection.assembly]::Load("mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies


	#region Form Objects

	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formChooseOU = New-Object 'System.Windows.Forms.Form'
	$cb_AdvancedFeatures = New-Object 'System.Windows.Forms.CheckBox'
	$Treeview = New-Object 'System.Windows.Forms.TreeView'
	$buttonOK = New-Object 'System.Windows.Forms.Button'
	$imagelist = New-Object 'System.Windows.Forms.ImageList'
	$ContextMenu = New-Object 'System.Windows.Forms.ContextMenuStrip'
	$changeDomainToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$newOUToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Form Objects

	#region Functions

	function Show-Error
	{
		Param([string]$Message)

		$msgbox = [System.Windows.Forms.MessageBox]::Show($Message,"Exception Report",0,16)
	}
	
    function Test-LDAPConnection
    {
	    Param(
		    $ComputerName
	    )
	    $TCPClient = New-Object Net.Sockets.TcpClient
	    $TCPClient.Connect($ComputerName,389)
	    $TestOutput = [pscustomobject]@{
		    ComputerName = $ComputerName
		    Connected = $TCPClient.Connected
		    IPAddress = ''
	    }
	    If ($TCPClient.Connected)
	    {
		    $TestOutput.IPAddress = $TCPClient.Client.RemoteEndPoint.Address.IPAddressToString
	    }
	    $TCPClient.Close()
	    $TestOutput
    }

	function Add-Node 
	{ 
		param ( 
			$RootNode, 
			$dname,
			$name,
			$Type,
			$HasChildren = $true
		)

		$newNode = new-object System.Windows.Forms.TreeNode
		$newNode.Name = $dname 
		$newNode.Text = $name
		If ($HasChildren)
		{
			$newNode.Nodes.Add('') | Out-Null
		}
		switch ($Type) {
			organizationalunit	{$newnode.ImageIndex = 3
								$newNode.SelectedImageIndex = 3}
			Domain 				{$newNode.ImageIndex = 2
								$newNode.SelectedImageIndex = 2}
			Default				{$newnode.ImageIndex = 0
								$newNode.SelectedImageIndex = 0}
		}
		$RootNode.Nodes.Add($newNode) | Out-Null 
		$newNode
	} 
	
	function Get-NextLevel
	{
	    param (
	        $RootNode,
	        $Type
	   	)
	   	
		If ($Type -eq 'Domain')
		{
			$ADObjects = $forest.domains | ?{$_.Name -eq $RootNode.Text} |
				select -ExpandProperty Children
			$RootNode.Nodes.Clear()
	        $ADObjects | % {
				$node = Add-Node -RootNode $RootNode -dname $_.GetDirectoryEntry().distinguishedName -name $_.name -Type $Type
	            Get-NextLevel -RootNode $node -Type $Type
	        }
		}
		else
		{
            If ($DomainIP)
            {
                If ($Credential)
                {
                    $ADsearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP/$($RootNode.Name)",$Credential.UserName,$Credential.GetNetworkCredential().password)
                }
                else
                {
                    $ADsearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP/$($RootNode.Name)")
                }
            }
            elseif ($Credential)
            {
                $ADsearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootNode.Name)",$Credential.UserName,$Credential.GetNetworkCredential().password)
            }
            else
            {
	    	    $ADsearcher.SearchRoot = "LDAP://$($RootNode.Name)"
			}
            $ADsearcher.filter = "(|(objectClass=organizationalUnit)(ObjectClass=container)(ObjectClass=builtinDomain))"
			$ADsearcher.SearchScope = 'OneLevel'
			
			IF ($cb_AdvancedFeatures.Checked)
			{	
				$ADObjects = $ADsearcher.FindAll()
			}
			else
			{
				$ADObjects = $ADsearcher.FindAll() | ?{$_.Properties['showinadvancedviewonly'][0] -eq $false -or 
					$_.Properties['showinadvancedviewonly'][0] -eq $null}
			}
		
		    If ($ADObjects) 
			{
		        $RootNode.Nodes.Clear()
				$ADObjects | % {
					$Type = $_.properties.objectclass | ?{$_ -ne 'top'}
                    If ($Credential)
                    {
                        $ADsearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry($_.Path,$Credential.UserName,$Credential.GetNetworkCredential().password)
                    }
                    else
                    {
					    $ADsearcher.SearchRoot = $_.Path
					}
                    If ($ADsearcher.FindOne())
					{
						Add-Node $RootNode $_.properties.distinguishedname[0] $_.properties.name[0] -Type $Type -HasChildren $true
					}
					else 
					{
						Add-Node $RootNode $_.properties.distinguishedname[0] $_.properties.name[0] -Type $Type -HasChildren $false
					}
		        }
		    }
		}
	}
	
	function Build-TreeView
	{ 
	    $treeNodes = $Treeview.Nodes[0]
	    	
		#Generate rootdomain node and add subdomain nodes
		If ($DomainDN)
		{
			$DomainName = $DomainDN.Replace(',DC=','.').Substring(3)
			$RootDomainNode = Add-Node -dname $DomainDN `
			-name $DomainName -RootNode $treeNodes -Type Domain
		}
		else
		{
		    $CurrentDomain = $Forest.Domains | ?{$_.Name -eq $env:USERDNSDOMAIN}
		    $Domain = $CurrentDomain.GetDirectoryEntry()
			$RootDomainNode = Add-Node -dname $Domain.distinguishedName `
				-name $CurrentDomain.Name -RootNode $treeNodes -Type Domain
		}
		#Copy the RootDomainNode to parent scope
		New-Variable -Name RootDomainNode -Value $RootDomainNode -Scope 1
		
		$treeNodes.Expand()
		$RootDomainNode.Expand()
	} 
	
	function Change-Domain
    {
		#region Form Objects
		
		$formBrowseForDomain = New-Object 'System.Windows.Forms.Form'
		$labelDomains = New-Object 'System.Windows.Forms.Label'
		$TreeviewForest = New-Object 'System.Windows.Forms.TreeView'
		$buttonCancelDomain = New-Object 'System.Windows.Forms.Button'
		$buttonOKDomain = New-Object 'System.Windows.Forms.Button'
        $buttonChangeForest = New-Object 'System.Windows.Forms.Button'
		$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	
		#endregion Form Objects
	
		#region Events
		
		$FormEvent_Shown={
			#Create domains treeview
			$DName = $forest.RootDomain.GetDirectoryEntry().distinguishedName
            $NodeProps = @{
                dname = $Dname
                name = $forest.RootDomain.Name
                RootNode = $TreeviewForest
                Type = 'Domain'
            }
			$RootDomainNode = Add-Node @NodeProps
		    Get-NextLevel -RootNode $RootDomainNode -Type Domain
			$TreeviewForest.ExpandAll()
		}
		
		$Form_StateCorrection_Load=
		{
			#Correct the initial state of the form to prevent the .Net maximized form issue
			$formBrowseForDomain.WindowState = $InitialFormWindowState
		}
		
		$Form_Cleanup_FormClosed=
		{
			#Remove all event handlers from the controls
			try
			{
				$formBrowseForDomain.remove_Load($FormEvent_Load)
				$formBrowseForDomain.remove_Load($Form_StateCorrection_Load)
				$formBrowseForDomain.remove_FormClosed($Form_Cleanup_FormClosed)
				$TreeviewForest.remove_DoubleClick($TreeviewForest_DoubleClick)
			}
			catch [Exception]
			{ }
		}
		
		$TreeviewForest_DoubleClick={
			#Click OK Button
			If ($TreeviewForest.SelectedNode.Nodes.Count -eq 0)
			{
				$buttonOKDomain.PerformClick()
			}
		}

        $buttonChangeForest_Click={
            $Refresh = $false
            New-Variable -Name ChangeForest -Value (Change-Forest) -Scope 1 -Force
            If ($ChangeForest.Forest)
            {
                $forest = $ChangeForest.Forest
                $Refresh = $true
            }
            If ($ChangeForest.Credential)
            {
                $Credential = $ChangeForest.Credential
                $Refresh = $true
            }
            If ($Refresh)
            {
                $TreeviewForest.Nodes.Clear()
                $TreeviewForest.Refresh()
                & $FormEvent_Shown
            }
	    }
	
		#endregion Events
	
		#region Form Code
		
		#
		# formBrowseForDomain
		#
		$formBrowseForDomain.Controls.Add($labelDomains)
		$formBrowseForDomain.Controls.Add($TreeviewForest)
		$formBrowseForDomain.Controls.Add($buttonCancelDomain)
		$formBrowseForDomain.Controls.Add($buttonOKDomain)
		$formBrowseForDomain.Controls.Add($buttonChangeForest)
		$formBrowseForDomain.AcceptButton = $buttonOKDomain
		$formBrowseForDomain.CancelButton = $buttonCancelDomain
		$formBrowseForDomain.ClientSize = '284, 279'
		$formBrowseForDomain.FormBorderStyle = 'FixedDialog'
		$formBrowseForDomain.MaximizeBox = $False
		$formBrowseForDomain.MinimizeBox = $False
		$formBrowseForDomain.Name = "formBrowseForDomain"
		$formBrowseForDomain.StartPosition = 'CenterScreen'
		$formBrowseForDomain.Text = "Browse for Domain"
		$formBrowseForDomain.add_Shown($FormEvent_Shown)
		#
		# labelDomains
		#
		$labelDomains.Location = '12, 9'
		$labelDomains.Name = "labelDomains"
		$labelDomains.Size = '99, 16'
		$labelDomains.TabIndex = 2
		$labelDomains.Text = "Domains:"
		#
		# TreeviewForest
		#
		$TreeviewForest.Location = '12, 28'
		$TreeviewForest.ImageList = $imagelist
		$TreeviewForest.ImageIndex = 2
		$TreeviewForest.Name = "TreeviewForest"
		$TreeviewForest.Size = '260, 193'
		$TreeviewForest.TabIndex = 1
		$TreeviewForest.add_DoubleClick($TreeviewForest_DoubleClick)
		#
		# buttonCancelDomain
		#
		$buttonCancelDomain.Anchor = 'Bottom, Right'
		$buttonCancelDomain.DialogResult = 'Cancel'
		$buttonCancelDomain.Location = '197, 244'
		$buttonCancelDomain.Name = "buttonCancelDomain"
		$buttonCancelDomain.Size = '75, 23'
		$buttonCancelDomain.TabIndex = 0
		$buttonCancelDomain.Text = "Can&cel"
		$buttonCancelDomain.UseVisualStyleBackColor = $True
		#
		# buttonOKDomain
		#
		$buttonOKDomain.Anchor = 'Bottom, Right'
		$buttonOKDomain.DialogResult = 'OK'
		$buttonOKDomain.Location = '116, 244'
		$buttonOKDomain.Name = "buttonOKDomain"
		$buttonOKDomain.Size = '75, 23'
		$buttonOKDomain.TabIndex = 0
		$buttonOKDomain.Text = "&OK"
		$buttonOKDomain.UseVisualStyleBackColor = $True
        #
		# buttonChangeForest
		#
		$buttonChangeForest.Anchor = 'Bottom, Left'
		#$buttonChangeForest.DialogResult = 'OK'
		$buttonChangeForest.Location = '13, 244'
		$buttonChangeForest.Name = "buttonChangeForest"
		$buttonChangeForest.Size = '91, 23'
		$buttonChangeForest.TabIndex = 0
		$buttonChangeForest.Text = "Change &Forest"
		$buttonChangeForest.UseVisualStyleBackColor = $True
        $buttonChangeForest.add_Click($buttonChangeForest_Click)
		#endregion Form Code
	
		#Save the initial state of the form
		$InitialFormWindowState = $formBrowseForDomain.WindowState
		#Init the OnLoad event to correct the initial state of the form
		$formBrowseForDomain.add_Load($Form_StateCorrection_Load)
		#Clean up the control events
		$formBrowseForDomain.add_FormClosed($Form_Cleanup_FormClosed)
		#Show the Form
		IF ($formBrowseForDomain.ShowDialog() -eq 'OK')
	    {
	        If ($ChangeForest)
            {
                $Props = @{
                    MemberType = 'NoteProperty'
                    Name = 'NewDomainDN'
                    Value = $TreeviewForest.SelectedNode.Name
                }
                $ChangeForest | Add-Member @Props -PassThru
            }
            else
            {
                [pscustomobject]@{
                    NewDomainDN = $TreeviewForest.SelectedNode.Name
                }
            }
	    }
	}
    
    function Change-Forest 
    {
	    [System.Windows.Forms.Application]::EnableVisualStyles()
	    $formChangeForest = New-Object 'System.Windows.Forms.Form'
	    $CB_UseCred = New-Object 'System.Windows.Forms.CheckBox'
	    $GB_Creds = New-Object 'System.Windows.Forms.GroupBox'
	    $MTB_Password = New-Object 'System.Windows.Forms.MaskedTextBox'
	    $TB_UserName = New-Object 'System.Windows.Forms.TextBox'
	    $labelUsernamePassword = New-Object 'System.Windows.Forms.Label'
	    $tb_DCName = New-Object 'System.Windows.Forms.TextBox'
	    $labelDomainControllerFQDN = New-Object 'System.Windows.Forms.Label'
	    $buttonOK = New-Object 'System.Windows.Forms.Button'
	    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

	    $CB_UseCred_CheckedChanged={
		    $GB_Creds.Enabled = $CB_UseCred.Checked
	    }
	
	    $buttonOK_Click={
		    $formChangeForest.Cursor = 'WaitCursor'	
		    $ServerName = $tb_DCName.Text
		    $UserName = $TB_UserName.Text
		    $Password = $MTB_Password.Text
		    try
		    {
			    $LDAPTest = Test-LDAPConnection -ComputerName $ServerName
                If ($LDAPTest.Connected)
                {
                    Write-Verbose "Connection successful."
                }
                else
                {
                    throw "Unable to connect to server '$ServerName' on LDAP port 389."
                }
                If ($CB_UseCred.Checked)
			    {
				    $DomainEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$ServerName",$UserName,$Password)
                    $SecPwd = $Password | ConvertTo-SecureString -AsPlainText -Force
                    $ForestCred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $SecPwd
			    }
			    else
			    {
				    $DomainEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$ServerName")
			    }
			    If ($DomainEntry.Name -eq $null)
			    {
				    throw "Unable to connect to server '$ServerName'"
			    }
			    $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain", $DomainEntry.name)
			    $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
			    $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $domain.forest)
			    $RemoteForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                $Props = @{
                    Server = $LDAPTest.IPAddress
                }
                If ($RemoteForest)
                {
                    $Props.Add('Forest',$RemoteForest)
                }
                If ($ForestCred)
                {
                    $Props.Add('Credential',$ForestCred)
                }
                New-Variable -Name ForestOutput -Scope 1 -Value ([pscustomobject]$Props)
		    }
		    catch
		    {
			    $Message = 'Connection failed! ' + $_.exception.message
			    $ErrorMsg = [System.Windows.Forms.MessageBox]::Show($Message,"Connection failure",0,16)
		    }
		    $formChangeForest.Cursor = 'Default'
	    }

	    $Form_StateCorrection_Load={
		    #Correct the initial state of the form to prevent the .Net maximized form issue
		    $formChangeForest.WindowState = $InitialFormWindowState
	    }
	
	    $Form_Cleanup_FormClosed={
		    #Remove all event handlers from the controls
		    try
		    {
			    $CB_UseCred.remove_CheckedChanged($CB_UseCred_CheckedChanged)
			    $buttonOK.remove_Click($buttonOK_Click)
			    $formChangeForest.remove_Load($Form_StateCorrection_Load)
			    $formChangeForest.remove_FormClosed($Form_Cleanup_FormClosed)
		    }
		    catch [Exception]
		    { }
	    }

	    # formChangeForest
	    #
	    $formChangeForest.Controls.Add($CB_UseCred)
	    $formChangeForest.Controls.Add($GB_Creds)
	    $formChangeForest.Controls.Add($tb_DCName)
	    $formChangeForest.Controls.Add($labelDomainControllerFQDN)
	    $formChangeForest.Controls.Add($buttonOK)
	    $formChangeForest.AcceptButton = $buttonOK
	    $formChangeForest.ClientSize = '253, 222'
	    $formChangeForest.FormBorderStyle = 'FixedDialog'
	    $formChangeForest.MaximizeBox = $False
	    $formChangeForest.MinimizeBox = $False
	    $formChangeForest.Name = "formChangeForest"
	    $formChangeForest.StartPosition = 'CenterScreen'
	    $formChangeForest.Text = "Change Forest"
	    #
	    # CB_UseCred
	    #
	    $CB_UseCred.Anchor = 'Bottom, Left'
	    $CB_UseCred.Location = '13, 68'
	    $CB_UseCred.Name = "CB_UseCred"
	    $CB_UseCred.Size = '172, 24'
	    $CB_UseCred.TabIndex = 3
	    $CB_UseCred.Text = "Use alternate credentials"
	    $CB_UseCred.UseVisualStyleBackColor = $True
	    $CB_UseCred.add_CheckedChanged($CB_UseCred_CheckedChanged)
	    #
	    # GB_Creds
	    #
	    $GB_Creds.Controls.Add($MTB_Password)
	    $GB_Creds.Controls.Add($TB_UserName)
	    $GB_Creds.Controls.Add($labelUsernamePassword)
	    $GB_Creds.Anchor = 'Left, Right'
	    $GB_Creds.Enabled = $False
	    $GB_Creds.Location = '13, 98'
	    $GB_Creds.Name = "GB_Creds"
	    $GB_Creds.Size = '227, 83'
	    $GB_Creds.TabIndex = 4
	    $GB_Creds.TabStop = $False
	    $GB_Creds.Text = "Alternate credentials"
	    #
	    # MTB_Password
	    #
	    $MTB_Password.Location = '85, 47'
	    $MTB_Password.Name = "MTB_Password"
	    $MTB_Password.PasswordChar = '*'
	    $MTB_Password.Size = '127, 20'
	    $MTB_Password.TabIndex = 5
	    #
	    # TB_UserName
	    #
	    $TB_UserName.Location = '85, 20'
	    $TB_UserName.Name = "TB_UserName"
	    $TB_UserName.Size = '127, 20'
	    $TB_UserName.TabIndex = 4
	    #
	    # labelUsernamePassword
	    #
	    $labelUsernamePassword.Location = '7, 20'
	    $labelUsernamePassword.Name = "labelUsernamePassword"
	    $labelUsernamePassword.Size = '71, 52'
	    $labelUsernamePassword.TabIndex = 0
	    $labelUsernamePassword.Text = "Username

Password"
	    #
	    # tb_DCName
	    #
	    $tb_DCName.Location = '13, 37'
	    $tb_DCName.Name = "tb_DCName"
	    $tb_DCName.Size = '227, 20'
	    $tb_DCName.TabIndex = 2
	    #
	    # labelDomainControllerFQDN
	    #
	    $labelDomainControllerFQDN.Location = '13, 13'
	    $labelDomainControllerFQDN.Name = "labelDomainControllerFQDN"
	    $labelDomainControllerFQDN.Size = '227, 16'
	    $labelDomainControllerFQDN.TabIndex = 1
	    $labelDomainControllerFQDN.Text = "Domain Controller (FQDN or IP address):"
	    #
	    # buttonOK
	    #
	    $buttonOK.Anchor = 'Bottom, Right'
	    $buttonOK.DialogResult = 'OK'
	    $buttonOK.Location = '166, 187'
	    $buttonOK.Name = "buttonOK"
	    $buttonOK.Size = '75, 23'
	    $buttonOK.TabIndex = 6
	    $buttonOK.Text = "&OK"
	    $buttonOK.UseVisualStyleBackColor = $True
	    $buttonOK.add_Click($buttonOK_Click)

	    #Save the initial state of the form
	    $InitialFormWindowState = $formChangeForest.WindowState
	    #Init the OnLoad event to correct the initial state of the form
	    $formChangeForest.add_Load($Form_StateCorrection_Load)
	    #Clean up the control events
	    $formChangeForest.add_FormClosed($Form_Cleanup_FormClosed)
	    #Show the Form
	    If ($formChangeForest.ShowDialog() -eq 'OK')
        {
            $ForestOutput
        }
    }

    #endregion Functions
	
    #region Script

	$FormEvent_Load={
		
		If ((gwmi win32_computersystem).partofdomain)
		{
			$forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			#Validate Domain variable if present
			If ($Domain)
			{
				If ($Domain -match 'CN=|OU=')
				{
					throw "'$Domain' is not a domain distinguished name."
					$formChooseOU.Close()
				}
				ElseIf ($Domain -match '\w+\.\w+')
				{
					If ($forest.Domains.Name -contains $Domain)
					{
						$DomainDN = "DC=$($Domain.ToUpper())" -replace '\.',',DC='
					}
					else
		            {
		                throw "No domain found with FQDN '$Domain'."
						$formChooseOU.Close()
		            }
				}
				ElseIf ($Domain -match 'DC=\w+,DC+')
				{
					If ([adsi]::exists("LDAP://$Domain"))
			        {
			            $DomainDN = $Domain
			        }
			        else
			        {
			            throw "'$Domain' does not exist in Active Directory."
						$formChooseOU.Close()
			        }
				}
				New-Variable -Name DomainDN -Value $DomainDN -Scope 1
			}
            If ($Credential)
            {
                $Credential = Get-Credential $Credential -Message "Please enter credentials to connect to domain '$Domain'"
			    try
			    {
				    $root = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainDN",$Credential.UserName,$Credential.GetNetworkCredential().password)
			    }
			    catch
			    {
				    Show-Error "Unable to connect to domain '$Domain'."
				    $host.SetShouldExit(1)
	    		    return
			    }
            }
            else
            {
			    $root = [ADSI]''
            }
            $ADSearcher = New-Object System.DirectoryServices.DirectorySearcher($root)
			New-Variable -Name ADSearcher -Value $ADSearcher -Scope 1
			New-Variable -Name Forest -Value $forest -Scope 1
		}
		elseif ($Domain)
		{
			If ($Domain -match 'CN=|OU=')
			{
				throw "'$Domain' is not a domain distinguished name."
				$formChooseOU.Close()
			}
			If ($Domain -match '\w+\.\w+')
			{
				$DomainDN = "DC=$($Domain.ToUpper())" -replace '\.',',DC='
			}
			ElseIf ($Domain -match 'DC=\w+,DC+')
			{
				$DomainDN = $Domain
                $Domain = $Domain.Replace('DC=','.').TrimStart('DC=')
			}
			New-Variable -Name DomainDN -Value $DomainDN -Scope 1
			$Credential = Get-Credential $Credential -Message "Please enter credentials to connect to domain '$Domain'"
			try
			{
                $LDAPTest = Test-LDAPConnection -ComputerName $Domain
                If ($LDAPTest.Connected)
                {
	                Write-Verbose "Connection successful."
                }
                else
                {
	                throw "Unable to connect to '$Domain' on LDAP port 389."
                }
                $DomainIP = $LDAPTest.IPAddress
				$root = New-Object -TypeName System.DirectoryServices.DirectoryEntry("LDAP://$DomainIP",$Credential.UserName,$Credential.GetNetworkCredential().password)
                New-Variable -Name DomainIP -Value $DomainIP -Scope 1
			}
			catch
			{
				Show-Error "Unable to connect to domain '$Domain'."
				$host.SetShouldExit(1)
	    		return
			}
            $ADSearcher = new-object System.DirectoryServices.DirectorySearcher($root)
		    New-Variable -Name ADSearcher -Value $ADSearcher -Scope 1
		}
        else
        {
            $Domain = Read-Host 'Please enter a domain to connect to'
            If ($Domain)
            {
                & $FormEvent_Load
            }
            else
            {
                return
            }
        }
        If ($Credential)
        {
            Set-Variable -Name Credential -Value $Credential -Scope 1 -Force
        }
		If ($MultiSelect)
		{
			$Treeview.CheckBoxes = $true
		}
        $cb_AdvancedFeatures.Checked = $AdvancedFeatures
	$ADSearcher.PropertiesToLoad.AddRange(@('name','distinguishedname','objectClass'))
	}
	
	$CreateOU=[System.Windows.Forms.NodeLabelEditEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.NodeLabelEditEventArgs]
		$NewNode = $_.Node
		$Label = $_.Label
		$NewNode.Name = "OU=$Label,$($NewNode.Parent.Name)"
		$objParent = [ADSI]"LDAP://$($NewNode.Parent.Name)"
		$objOU = $objParent.Create("organizationalUnit", "ou=$Label")
		try
		{
			Write-Verbose "Creating OU '$($_.Label)'."
			$objOU.SetInfo()
			
			Write-Verbose "Setting 'protect from accidental deletion'."
		    $ObjectSecurity = $objOU.psbase.ObjectSecurity
		    $EveryOne = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList '', 'Everyone'
		    $Act = [System.Security.AccessControl.AccessControlType]::Deny
		    $ADRights = [System.DirectoryServices.ActiveDirectoryRights]::Delete
		    $NewRule = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule -ArgumentList $EveryOne, $ADRights, $Act
		    $ObjectSecurity.AddAccessRule($NewRule)
	        $ADRights = [System.DirectoryServices.ActiveDirectoryRights]::DeleteTree
		    $NewRule = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule -ArgumentList $EveryOne, $ADRights, $Act
		    $ObjectSecurity.AddAccessRule($NewRule)
		    $objOU.psbase.CommitChanges()
		}
		catch [System.UnauthorizedAccessException]
		{
			$TryWithCred = $true
		}
		catch
		{
			$Treeview.SelectedNode = $NewNode.Parent
			$Treeview.Nodes.Remove($NewNode)
			Show-Error -Message $_.Exception.GetBaseException().Message
		}
		If ($TryWithCred)
		{
			try
			{
				If (!$ADWriteCred)
				{
					$Cred = $host.ui.PromptForCredential("Insufficient rights in AD detected", "Please enter credentials with sufficient rights in Active Directory.", "", "NetBiosUserName")
				
					If ($Cred)
					{
						#Create variable in parent scope
						New-Variable -Name ADWriteCred -Value $Cred -Scope 1
					}
					else
					{
						$Treeview.SelectedNode = $NewNode.Parent
						$Treeview.Nodes.Remove($NewNode)
						return
					}
				}
				$DomainDN = ($NewNode.Name.Split(',') | ?{$_ -like 'DC=*'}) -Join ','
				$objParent = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($NewNode.Parent.Name)", $($ADWriteCred.UserName),$($ADWriteCred.GetNetworkCredential().password)
				$objOU = $objParent.Create("organizationalUnit", "ou=$Label")
	
				Write-Verbose "Creating OU '$Label'."
				$objOU.SetInfo()
				Write-Verbose "Setting 'protect from accidental deletion'."
			    $ObjectSecurity = $objOU.psbase.ObjectSecurity
			    $EveryOne = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList '', 'Everyone'
			    $Act = [System.Security.AccessControl.AccessControlType]::Deny
			    $ADRights = [System.DirectoryServices.ActiveDirectoryRights]::Delete
			    $NewRule = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule -ArgumentList $EveryOne, $ADRights, $Act
			    $ObjectSecurity.AddAccessRule($NewRule)
		        $ADRights = [System.DirectoryServices.ActiveDirectoryRights]::DeleteTree
			    $NewRule = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule -ArgumentList $EveryOne, $ADRights, $Act
			    $ObjectSecurity.AddAccessRule($NewRule)
				$ADRights = [System.DirectoryServices.ActiveDirectoryRights]::DeleteChild
			    $NewRule = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule -ArgumentList $EveryOne, $ADRights, $Act
			    $ObjectSecurity.AddAccessRule($NewRule)
			    $objOU.psbase.CommitChanges()
			}
			catch
			{
				$Treeview.SelectedNode = $NewNode.Parent
				$Treeview.Nodes.Remove($NewNode)
				$Text = $_.Exception.Message
				Show-Error -Message $_.Exception.GetBaseException().Message
			}	
		}
		$Treeview.LabelEdit = $false
	}
	
	$buttonOK_Click={
		If ($Treeview.SelectedNode -and !$MultiSelect)
		{
			New-Variable -Scope 1 -Name SelectedObject -Value ([pscustomobject]@{
				Name = $Treeview.SelectedNode.Text
				DistinguishedName = $Treeview.SelectedNode.Name
			})
		}
	}
	
	$formChooseOU_Shown={
		#Build treeview when form is shown
		$formChooseOU.Cursor = 'WaitCursor'
		try
		{
			Build-TreeView
		}
		catch
		{
			Show-Error ($_ | Out-String)
		}
		finally
		{
			$formChooseOU.Cursor = 'Default'
            $formChooseOU.BringToFront()
		}
	}
	
	$Treeview_BeforeExpand=[System.Windows.Forms.TreeViewCancelEventHandler]{
		#Get next level for current node
        If ($_.Node.Level -eq 1 -and $RootOU)
        {
            $_.Node.Nodes.Clear()
            $RootNode = Add-Node -dname $RootOU `
				-name $RootOU.Split(',')[0].Substring(3) -RootNode $_.Node
            $RootNode.Expand()
        }
        elseIf ($_.Node.Level -ne 0)
		{
			Get-NextLevel -RootNode $_.Node
		}
	}
	
	$cb_AdvancedFeatures_CheckedChanged={
		#Refresh the treeview
		If ($Treeview.Nodes[0].Nodes)
		{
			$Treeview.Nodes[0].Nodes.Clear()
			& $formChooseOU_Shown
		}
	}
	
	$Treeview_DoubleClick={
		#Click OK Button
		If ($Treeview.SelectedNode.Nodes.Count -eq 0)
		{
			$buttonOK.PerformClick()
			New-Variable -Scope 1 -Name SelectedObject -Value $SelectedObject
		}
	}
	
	$changeDomainToolStripMenuItem_Click={
		#call Change Domain form
        $Refresh = $false
		$ChangeDomain = Change-Domain
		If ($ChangeDomain.NewDomainDN)
		{
			New-Variable -Name DomainDN -Value $ChangeDomain.NewDomainDN -Scope 1 -Force
			$Refresh = $true
		}
        else
        {
            return
        }
        If ($ChangeDomain.Forest)
        {
            Set-Variable -Name Forest -Value $ChangeDomain.forest -Scope 1 -Force
            New-Variable -Name DomainIP -Value $ChangeDomain.server -Scope 1 -Force
			$Refresh = $true
        }
        If ($ChangeDomain.Credential)
        {
            Set-Variable -Name Credential -Value $ChangeDomain.Credential -Scope 1 -Force
			$Refresh = $true
        }
        If ($Refresh)
        {
            $Treeview.Nodes[0].Nodes.Clear()
		    $Treeview.Refresh()
		    Build-TreeView
        }
	}
	
	$newOUToolStripMenuItem_Click={
		#Create new node and edit the label
		$SelectedObject = $Treeview.SelectedNode
		If(!$SelectedObject.IsExpanded){
	        $SelectedObject.Expand()
	    }
		$newOuNode = new-object System.Windows.Forms.TreeNode 
	    $newOuNode.text = "New OU" 
	    $newOuNode.Name = "New OU Name" 
		$newOuNode.ImageIndex = 3
		$newOuNode.SelectedImageIndex = 3
	    $SelectedObject.Nodes.Add($newOuNode) | Out-Null 
	    $Treeview.SelectedNode = $NewOUNode
	    $Treeview.LabelEdit = $true
	    $newOuNode.BeginEdit()
	}
	
	$Treeview_NodeMouseClick=[System.Windows.Forms.TreeNodeMouseClickEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.TreeNodeMouseClickEventArgs]
		if($_.Button -eq 'Right')
	    {
			$Treeview.SelectedNode = $_.Node
			if ($_.Node.ImageIndex -eq 0)
			{
				return
			}
			$ContextMenu.Items.Clear()
			if ($_.Node.ImageIndex -lt 3 -and (gwmi win32_computersystem).partofdomain)
			{
				$ContextMenu.Items.Add($changeDomainToolStripMenuItem)
			}
			if ($_.Node.ImageIndex -gt 1 -and !$HideNewOUFeature)
			{
				$ContextMenu.Items.Add($newOUToolStripMenuItem)
			}
			If ($ContextMenu.Items.Count -gt 0)
		    {
		        $ContextMenu.Show($Treeview, $_.Location)
		    }
		}
	}
	
	$Treeview_BeforeCheck=[System.Windows.Forms.TreeViewCancelEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.TreeViewCancelEventArgs]
		#Prevent checking root node
		if($_.Node.ImageIndex -eq 1)
	    {
	        $_.Cancel = $true
	    }
	}
	
	$Treeview_AfterCheck=[System.Windows.Forms.TreeViewEventHandler]{
	#Event Argument: $_ = [System.Windows.Forms.TreeViewEventArgs]
		If (!$SelectedObject)
		{
			New-Variable -Name SelectedObject -Scope 1 -Value (New-Object collections.arraylist)
		}
		If ($_.Node.Checked)
		{
			$SelectedObject.Add([pscustomobject]@{
				Name = $_.Node.Text
				DistinguishedName = $_.Node.Name
			})
		}
		else
		{
			$DN = $_.Node.Name
			$SelectedObject | %{
				If ($_.DistinguishedName -eq $DN)
				{
					$Remove = $_
				}
			}
			$SelectedObject.Remove($Remove)
		}
	}
	
	#endregion Script
	
	#region Events
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$formChooseOU.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$cb_AdvancedFeatures.remove_CheckedChanged($cb_AdvancedFeatures_CheckedChanged)
			$Treeview.remove_AfterLabelEdit($CreateOU)
			$Treeview.remove_BeforeCheck($Treeview_BeforeCheck)
			$Treeview.remove_AfterCheck($Treeview_AfterCheck)
			$Treeview.remove_BeforeExpand($Treeview_BeforeExpand)
			$Treeview.remove_NodeMouseClick($Treeview_NodeMouseClick)
			$Treeview.remove_DoubleClick($Treeview_DoubleClick)
			$buttonOK.remove_Click($buttonOK_Click)
			$formChooseOU.remove_Load($FormEvent_Load)
			$formChooseOU.remove_Shown($formChooseOU_Shown)
			$changeDomainToolStripMenuItem.remove_Click($changeDomainToolStripMenuItem_Click)
			$newOUToolStripMenuItem.remove_Click($newOUToolStripMenuItem_Click)
			$formChooseOU.remove_Load($Form_StateCorrection_Load)
			$formChooseOU.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Events

	#region Form Code
	
	#
	# formChooseOU
	#
	$formChooseOU.Controls.Add($cb_AdvancedFeatures)
	$formChooseOU.Controls.Add($Treeview)
	$formChooseOU.Controls.Add($buttonOK)
	$formChooseOU.AcceptButton = $buttonOK
	$formChooseOU.ClientSize = '342, 535'
	$formChooseOU.FormBorderStyle = 'FixedDialog'
	#region Binary Data
	$formChooseOU.Icon = [System.Convert]::FromBase64String('
AAABAAkAICAQAAEABADoAgAAlgAAABgYEAABAAQA6AEAAH4DAAAQEBAAAQAEACgBAABmBQAAICAA
AAEACACoCAAAjgYAABgYAAABAAgAyAYAADYPAAAQEAAAAQAIAGgFAAD+FQAAICAAAAEAIACoEAAA
ZhsAABgYAAABACAAiAkAAA4sAAAQEAAAAQAgAGgEAACWNQAAKAAAACAAAABAAAAAAQAEAAAAAAAA
AgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAIAAAACAgACAAAAAgACAAICAAACAgIAAwMDAAAAA
/wAA/wAAAP//AP8AAAD/AP8A//8AAP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAiIAAAAAAAAAAAAAAAACIiIiHMAAAAAAAAAAAAAiI//iIiHcAAAAAAAAAAIiP//+IiIiH
dwAAAAAAiI//////iIiIiHhwAAAAAIj///////iIiIiHeHAAAACIj////4iIiIiIiIiHgAAAiIj/
iIiIiIj/iIiHeIcAAIiIiIiIiIiIiP+IiIeIcACIiHd3d3d3eIeI+IiHeIhwczMRMzd3czeIeI/4
iIiIcAczN3iIiIczeId4j4iHiHAACIiP///4i7OIh4iPiHiAAAAIiI//iIu7M4h3iP+HcAAAAAiI
iIi7u7u3iHiIiHAAAAAACIeP///4i7iHeIgAAAAAAAAAiI+IiLu7t3AAAAAAAAAAAACDMzMzMzM3
AAAAAAAAAAAAAAAAAAAACIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////////////////////////////////////
///////8f///wB///gAP//AAA/8AAAH/AAAAfwAAAB8AAAAPAAAABwAAAAEAAAABgAAAAeAAAAH4
AAAB/gAAAf+AAAP/8AAf//wAD////+f/////////////////////KAAAABgAAAAwAAAAAQAEAAAA
AAAgAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAIAAAACAgACAAAAAgACAAICAAACAgIAAwMDA
AAAA/wAA/wAAAP//AP8AAAD/AP8A//8AAP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIiIAAAAAAAA
AAAIiIiIcAAAAAAAAIiI/4iIhwAAAAAIiP//+IiIh3gAAACI/////4iHiHeAAACIj//4iIiIeIiI
AACIiIiIiIiPiIh4iACIiId4iIiI+IiHiHCHczMzd3eIiPiHeIeDAzd4h3M3iI+Id4gAiIj//4iz
eIiPh4gAAIiIiIu7M3iIiHcAAACIiIiIizeIiIcAAAAAiI//iLt3gAAAAAAAAIhzMzMzMAAAAAAA
AAAAAAAAdwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A////AP///wD///8A////AP///wD/
w/8A/gH/APAA/wCAAD8AAAAfAAAADwAAAAMAAAABAAAAAAAAAAAAwAAAAPAAAAD8AAAA/wAHAP/A
BwD///MA////AP///wAoAAAAEAAAACAAAAABAAQAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAgAAAgAAAAICAAIAAAACAAIAAgIAAAICAgADAwMAAAAD/AAD/AAAA//8A/wAAAP8A/wD//wAA
////AAAzMzMzMzMAAHd3d3MzMwAAd3d3czMzAACPiIiIiIMAAI+IiIiIgwAAj4iIiIiDAACPiIiI
iIMAAI+IiIiIgwAAj4iIh3eDAACPj///94MAAI+P///3gwAAj4////eDAACPiIiId4MAAHd3d3Mz
MwAAd3d3d3d3AAAAAAAAAAAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMA
AMADAADAAwAAwAMAAMADAADAAwAA//8AACgAAAAgAAAAQAAAAAEACAAAAAAAAAQAAAAAAAAAAAAA
AAEAAAABAAAAAAAAEj1BABlFSgAZSlIAIUpNAC5dXgAwV1kAMWRtAD9jaQA8aXoAVWttAFhpaQBK
bnQATH5/AF16ewAAaZMACmuTAAl0mgA/eIIAHXqgACp+owBNd4EAS32LAGF9gAAShJgALouZADyG
mQAKh6kAE4ujACWNqAA/jaMAMourAC6SrgAumbIAOJayAEmDjABEgpMAS4eQAFCHkABSiZIAVZOd
AFmUngBhh4kAaIeMAGmMjQBnlZsAaZmbAHWUkwB7lJMAcJyfAHWdnwBQkKQAX5ujAEKfugBqnaEA
UKK7AF2puwBpoaUAYqSqAGmlqQBvqq4AfaOsAHWprQB4q68Aaqm4AGqtvQB0rrIAea6xAHKxtgB5
s7cAdbS4AHy1uQBzu74Adbm+AH64vAAXo8kAAKbUABSy1QAdtdkAHrvbACGjyQA4uM4ALL3eABi7
4ABdssEATLbQAGOqwABmqsAAZq/AAGGtxQBqrcIAbrDBAHW8wQB5u8EAf73CAHy0yAB4v8oAZLfQ
ADrB3wAp0t8APNDaAD7N5wA23OQAQsPfAEvO2AB5wcYAf8DEAHzCygBuw9YAdsTUAGPQ2QB92NwA
RsTgAE3J4QBH3OsAUtDqAF7b7QBh1ecAWPz/AH3k8QB7//8AgKmqAImtsQCCsrQAhLa9AI67uwCR
ubwAmMC+AIe+xgCLvMQAgrjKAJy/xgCQv8oAgL7SAIi90ACCwcYAi8PHAITFzACNxMgAi8PNAIjH
zACEys8AiMnOAJTAwQCXxsgAm8bIAJvIzACGzdIAhs3UAI7J0ACIzNAAjs3SAJjD0QCTyNQAh9DU
AIvQ0wCI0tYAjdHVAIvS2ACP1NoAjNnfAJHS1gCd09QAk9XZAJvV2gCT2NoAlNjaAJHY3ACU2NwA
os3ZALLO0gCk09YArNHTAK7W2QCs298AstPWALHV2wCO2uAAk9vgAJXb4ACV3OEAl9/kAJzd5ACq
0eAAq9XgAKLZ4wCg3+YAqN7hAKjY5ACu2uUAqdznALDT4QC11eMAs9nlAJng5ACd4ucAneTnAJ3h
6gCe5uoAn+fsAI/s8wCd8vsAm///AKrg5ACi5+wAquToAKzl6wCl6e4AtOPpALrk6wC26+4At+zu
ALzq7QCk6/AAquvwAKjt8QCx7fEAuO3wAKrw8gCq8PYArPL3AK319gCm8fgArvP4ALXw8wC28/YA
vfHzALvy9gC09foAs/r+ALz//wDI19wAw9zoAMHq7QDF6e8AwuzuAN7u7wDD7fAAxu3wAMju8ADC
8fMAx/HzAMjx8gDJ8vUAzPP2AM329gDP+PcAxff6AMP8/wDK/PwA2PD0ANH//wDY//8A6P//APL/
/wD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3uW0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6MC3ztfa
pS0JAAAAAAAAAAAAAAAAAAAAAAAAAADAwdHu8u/Z19eTjHlBAAAAAAAAAAAAAAAAAAAAAMC96fL2
8vLu6dLX15OaoC8rWQAAAAAAAAAAAADAvb3u8vLy8vLy8u7p2Nfak5OcXjJ9WwAAAAAAAAAAAMXa
4fDy8vLy8vLy8uLXw7NIaZOeoBcwlVkAAAAAAAAAw9fX4fDw8vLy8tTLpZyew8WTXImeh0V7j0GG
AAAAAADD2tfa5OLS0L2koaXD19rf5eXFa1ycsQ4sooFZAAAAALna18WzpZyJaYeJiYeHh4eVz+Xk
pVxrlUdDjI9BAAAAw7OTazolIyYnJygpNDtCSkVCw9/l14lpjQs+nH98WgBBEgcDAQEEBggMFRYk
Gh44rZA5ic/l5bNpRkpHjo8/AAAnAwIFDS57kaytq4lUIRs1qa8+PKXa5dqcSApHnkQAAAAAgWCN
zvn7+/n45Ml3c05LYbB+NofM3+XFXkSHSAAAAAAAANBrgK739+TeyHd0ZVNMUIWqMjyl1+Xlnio7
AAAAAAAAAADQjHyk13d1cXBnYlJPTVWYejaHw9fapToAAAAAAAAAAAAA6YlZmfr+/v385sp4dnJs
g3lCiJKsAAAAAAAAAAAAAAAAAACJhOfs1MNvbmhkZmNRNz0AAAAAAAAAAAAAAAAAAAAAAAAAYBkY
HB0gIh8UExEPEDMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX4IAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAP/////////////////////////////////////////////////8f///wB///gAP//AAA/8A
AAH/AAAAfwAAAB8AAAAPAAAABwAAAAEAAAABgAAAAeAAAAH4AAAB/gAAAf+AAAP/8AAf//wAD///
/+f/////////////////////KAAAABgAAAAwAAAAAQAIAAAAAABAAgAAAAAAAAAAAAAAAQAAAAEA
AAAAAAAKNjwAH0tOACFQVgAsWV8AIFdhAClcZAA4ZGwAQGxuAEBtdQARaIMAMnePAEJ0gQBFf4YA
QnuJAAuBnwAcgp0AH5CuADOMpAA5i6EAKJSvAC2arAAsmbEAIZu7ADuuvgBEgo8AXoqMAFCQmwBV
k5wAX5CfAF+WnABjhIUAYYeKAG2XmQBmmJsAeZaVAHaZmgBIl6cARp+rAEafrABYkKEAWZWlAGOZ
pwBgnqoAaKGnAGyqrwB/oKEAeqKlAHSnqwB7o6wAe6WvAHCorQB3q68Abq61AHajswBxqrIAdKyy
AHevtQByrrgAd6+8AG+yuQB9srYAcLO5AHO0ugB5tLkAeLu/AH64vQACmsUAErjdACi52gBur8AA
d7fDAHe9wgB3vsQAe7rAAHm9wwB8uMQAfL/GAHq5zAB6vtEAM8XOADzM5gA80uYAfsHFAH7DyAB8
wcwAf8TOAF/a6wBj2u4AaNnqAHra7QBV9PUAfebwAIyvsQCArr0AgLK2AIqxsgCJsrcAkLW6AJO4
vwCMvsAAlr7CAI7CwgCBxcoAh8fJAIPHzACKx8wAg8rOAIvKzwCMys0AmMLFAJ7GxQCRxMkAnsLM
AJrFzACSyc0Als3OAJ3JyACCwNUAgsrRAITL0ACHzNcAi83SAI3N0QCLztQAjM7bAJHP0QCZyNIA
n8nRAJvJ1ACO0NUAj9PfAIzX3QCe0NAAkdXaAJTX3QCd1NgAn9bfAJTZ2wCR2NwAlNndAKLI0gCj
ytUApcrUAKzJ0wCwy9YApNDXAKbW2wCq1dwAu9beAIXV4gCB2esAhNrrAIrc6gCa1uAAl9vgAJbc
4QCZ3+MAm9vlAJjf5ACk3OAAqt3jALXd4wCb4+cAnOLlAJ7k6QCY7vcAgfPzAKfh5QCg4ukAp+fr
AKbm7ACq4eoApOrvAKft7wCq7O8AtubqALDj7QC44usAueXoALbq6wC37O8Av+nrALvp7wCm6/EA
pOzwAKju8gCq7/QAs+zwALnv9ACu8PMAq/H0AKzx9QCt9/UAo/D4AK7z+QCv9PgAsvD0ALj3/ACy
+f0Atfr9ALX8/wDK4OYAxOvuAMXt8QDB8PIAwvP0AMT09QDJ8PEAzPLzAMny9QDM8/cAyPb2AM32
9gDO+PgAyv3/ANP9+gDT//8A7vHyAPf7+wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlYF5XgAAAAAAAAAAAAAAAAAA
AAAAkXGBoK6kQh0AAAAAAAAAAAAAAAAAkI2UttHRtbuLQmYqAAAAAAAAAAAAjY2izNHW0dHLvMCL
QGsjOXYAAAAAAACewM/R0dHR0dHPqptCM1SEJC5PAAAAAAClwMXP0dHMs5NtaXybaTx5aT1gTAAA
AACqwLuvqIhzanybrrvDyZ4/SHkfZHBOAACqu5tMPzU1P0xTaX6LpcnDUz9rIV91RgBHHA0GAwQH
CQwOGRxfc4vAyYtIIjNtbzqSBQECCBouXWA5JRATMW98pMm4aS0ggkIAAHFUh7XR2dbGpllFFyli
cIe7yYssQEwAAAAAtpqhsL3CXFdRREMSMWVtnru4MB0AAAAAAACziISWmZiYWlhSGCtjhXx8bTkA
AAAAAAAAALB9ytrb2cCnW1AmOoEAAAAAAAAAAAAAAAAArFQmFRYWFBEPCgsAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAADYpAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAD///8A////AP///wD///8A////AP///wD/w/8A/gH/APAA/wCAAD8AAAAfAAAADwAAAAMA
AAABAAAAAAAAAAAAwAAAAPAAAAD8AAAA/wAHAP/ABwD///MA////AP///wAoAAAAEAAAACAAAAAB
AAgAAAAAAAABAAAAAAAAAAAAAAABAAAAAQAAAAAAAABfggAljaMALZClADaTpgA/l6gAQouhAEma
qgBTnqsAXKKtAGalrwBvqbEAeKyyADOixQBBrMsAQ63MAEWvzQBGsM4ASbLPAFC10QBZudMAW7zV
AGe/1wBowtkAdsXbAHfI3QBo0OYAcNPoAHnW6QCGzN8Ahc7kAIjP4QCC2usAjN3tAJnT4wCb1uUA
n9XlAKTY5wCq2+gAr93qAJbh7gCf5fAAqejyALLs9AC77/UAvP//AMX//wDP//8A2f//AOL//wDs
//8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAA////AAAAAQEBAQEBAQEBAQEBAAAAAAYMCwoJCAcFBAMCAQAA
AAAGDAsKCQgHBQQDAgEAAAAAHiwrKikoISAcGxoNAAAAAB4sKyopKCEgHBsaDQAAAAAeLCsqKSgh
IBwbGg0AAAAAHiwrKikoISAcGxoNAAAAAB4sKyopKCEgHBsaDQAAAAAeLCIdGBYUExEOGg0AAAAA
HiwiMjEwLy4tDhoNAAAAAB4sJTIxMC8uLQ4aDQAAAAAeLCYyMTAvLi0RGg0AAAAAHiwmIh8ZFxQT
EhoNAAAAAAYMCwoJCAcFBAMCAQAAAAAGBgYGBgYGBgYGBgYAAAAAAAAAAAAAAAAAAAAAAADAAwAA
wAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAAD/
/wAAKAAAACAAAABAAAAAAQAgAAAAAACAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdKbGAYG0zTSOwtRUhLzS
r5zX4e1qv9HqCF59egAnQg4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgbHMQI2+1IaY
xtiyqNTf9Kzl6/+p7vH/qvHy/5PY2v9nlZv/ADtSww9tjTMAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaJ/CA4y91CR/tM5JjL3U
pafU4dy45Ov3xu3x/8rx8v/I7vD/uO3w/6nt8f+o7PH/hs7T/4jHzP+Aqqr/YKq67wd5oHsAR2YK
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbaLBPoa4z26R
wtasqtjk8cDq7v/K8/X/z/j3/8z29v/J8vP/xu3w/8Pr7f+26+7/qOzw/6jt8f+HztL/h9DU/43a
3/91lJP/aIeM/yqRsb0AX4s3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf7jO
nJLP3sul2+bzw+3w/8329v/O9vf/zPP2/8ry9f/I8vX/yfL1/8jx8v/H7vH/wuzu/7Ht8f+p7vL/
qvDz/4bN0P+HztP/idLW/3+9wv91nZ//jru7/2Orve0Dc5t7AE9zCQAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAACd4er/rvX2/7bz9v/D8vP/yfPz/8jy9f/I8fP/yfP1/8ny9f/J8vX/yfHz/8jw
8v+98fP/quvw/53i5/+V2+D/c7u+/3nBxv+HztP/i9PY/4zZ3/9hfYD/e5ST/47J0P81kK3BAGiQ
NgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJvf5/mp7vL/qu7z/7Xw8//B8fP/x/Hz/8rz9v/M8/f/
yvP2/8jy8/++6+7/quDk/5TY2v+M09b/j9Xb/5ng5f+f5+z/h87U/3S7wf+Ax87/i9LY/4LCx/95
srf/gbK1/5XBwv9ep7rqFH6kgQAAAAMAAAAAAAAAAAAAAAAAAAAAm97m+anw8v+p7vP/rPL3/7X2
+P+78vb/t+zu/7Tm6v+o3uH/mdXZ/5HS1v+T2Nv/neTn/6bt8v+q8Pb/r/P4/7L4/v+1/P//n+br
/3zDyP93vsP/iNLW/47a4P9denv/aYyN/53T1P+Mu8H/Mo+swQBqkzcAAAAAAAAAAAAAAACe3+b5
q/Dz/6jt8v+f5ur/l9/k/5LY3f+L0NP/hMbK/3/AxP+Dwsb/h8bL/4bHyv+ExMf/hMLG/4PAxf+F
wsb/js3S/6bp7v+0+///sPf8/5HY3P90u8H/fcbM/4jM0P98tbj/ea6x/43EyP+XwMD/Y6i58QB1
n38AVnUJAAAAAJ7e5fqT2+D/hs3U/3rCyf9hoqr/S4eQ/0mDjP9Qh5D/UIiS/1OMlf9Vk53/WZSe
/1+bo/9ppKj/dK6y/365vP95tLf/dK+0/5rc4f+v8/j/t/7//6Xs8P+Axsz/ecHH/4PMz/9YaWn/
dqqu/47Q1f+YwL7/h7e9/yuLqbMAAAAKaqm4/z94gv8xZG3/G0tT/xA7Qf8VP0L/IUpN/zBXWf8/
Y2n/Sm50/013gf9LfYv/RIKT/zyGmf8/jaP/Xam7/67W2/+Xxsj/aaGl/4XDyP+k6u7/sfj9/7L5
/v+V3OH/ecLH/3W0uP9/uLz/fLa7/4jJzv+SwML/d6uv/QAAACqCy9NpPHiA3hhKUf8ZRUr/Ll1e
/0x+f/9pmZv/g7K0/5vGyP+r0NL/rtbY/6TT1v+Hxs3/XbLB/y6Zsv8Kh6n/Qp+6/6LN2f+y09b/
dais/2+prv+U2Nz/qvD1/7X7//+r8vf/iNLW/3S6vv9Va23/fLW4/4/V2/9ur7X2AAAAKAAAAACm
7fMEh8zVZGeqtsJ5v8n/hsjP/6rk6P/J+vr/0P///9L////M/v//w/z//7T3/v+d8vv/f+bz/1LQ
6v8dtdn/F6PJ/2S30P+x1dv/kbm8/2icof+AwMT/oufs/671+P+1+///nufr/3m7wf9xsbb/hcLH
/3O3vPZNTU0qAAAAAAAAAAAAAAAAktbfAqLq9UyQ1d6tf8DL/4e+xv+s29//x/f5/8P3+/+38/n/
pvH4/4/s8/955fD/Xtvt/z7N5/8Yu+D/AKbU/yGjyf+AvtL/ss7S/3Ccn/9vq67/k9fa/6br8P+0
+f7/svz//4/U2P9hh4n/ZqSp+MHBwUEAAAAAAAAAAAAAAAAAAAAAAAAAAKvt8wKw8PUuktPeqIK+
yeuCtr3/ntXc/6Do8f9/4vD/YdXn/03J4f9GxOD/QsPf/zrB3/8svd7/Hrvb/xSy1f9MttD/mMPR
/4mtsf9snqH/gMLG/5rg5P+m7PH/rPX3/5PU2P9kpqv/x8fHHAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAn+LtKI7U4YF4v8voZq/A/5PI1P/Y8PT/8v////P////o////2P///7z///+b
////e////1j8//9H3Ov/bsPW/5y/xv9/p6r8b6qv83O3vNJorLKoY6mvhF+lq3MAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJLi7B6C1uFpc7rI3pC/yv/I19z/3u7v
/7vq7P+Z4OT/fdjc/2PQ2f9Lztj/PNDa/zbc5P8p0t//OLjO/0yguvg6c4GnRHB1aDxiZwsAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAWsbXZUuqur8ui5n/EoSY/wmGn/QDe5vcAHud0QB5nccAb5fNAGWR1QBplOIAb5f2AGmT/wpr
k/8pd5DQMm5+eAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAVdLfA0CxxkQmjq5FKIuwNSiGriIukrgaM5vBEiqTuhYig7Ad
I4ayJwxxoDcOdKRBIIWvWBp8n5Ecf6CMHoGlKgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/////////////////////////////
////////////////////4B///4AP//AAA//AAAH/AAAAfwAAAD8AAAAPAAAABwAAAAEAAAAAAAAA
AAAAAACAAAAA4AAAAPgAAAD/AAAB/8AAB//4AAf//AAD/////////////////////ygAAAAYAAAA
MAAAAAEAIAAAAAAAYAkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABCcHgCQnB4BQAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAgbLNAYCyzRFzpLpLfLC/g4e+zNOBytXyIXGMkQBFbBMAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGuesySBr7xSe6m7mIe0wc6d
yND4pNzg/6ft7/+c4uX/fri9/ydpfbwUmMMxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAQnB4CF+QozJmma1of628pY68ydCq1dz+v+nr/8zy8//K8PH/t+zv/6ru8/+R2N3/fri9
/47Cwv9Rjp3kE4OvYgB2rgEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABso7R7erHAq5PBztey3OL0
xOzw/8/29//O+Pj/zPX1/8jw8v/E6+7/s+zw/6zy9v+U2t3/erS5/4PKzf95lpX/da60+jeavJ8A
dqoaAAAAAAAAAAAAAAAAAAAAAAAAAACZ2+X4rvDz/8T09f/O9/b/y/P2/8ry9f/J8vX/yvL1/8rw
8v/B8PL/p+fr/5fb4P94u7//cKit/37DyP+M193/dpma/3+gof9csMfPB4GwRAAAAAAAAAAAAAAA
AAAAAACf5Or+q/H0/7Lw9P/C8/T/yPP1/8zz9//G7/L/ueXo/6bW2/+Mys7/g8fM/4vO1P+W3OH/
gcXL/2+yuf+CytH/g8TJ/32ytv+KsbL/c7PA7iqUuH4gj7QJAAAAAAAAAACf4un6rfL1/6nv9P+q
7O//p+Hl/53U2P+Syc3/isfM/4zN0f+X3eD/pOrv/6nt8v+u8/n/tfz//5nf4/9ztLr/d73C/4TL
0P9jhIX/jL7A/5LEyf88mbWtAYCxKgAAAACl5uz6qO/z/5bd4/98v8b/cLO5/2+vtv9trrX/c7S7
/3m9w/9+wMT/h8fJ/5HP0f+U2dv/neTp/7X6/f+v9Pj/f8LH/3Czuv+Dy8//bZeZ/4Gztv+dycj/
WaS33hOIslF3t8P/VZOc/0V/hv8pXGT/IVBW/yxZX/84ZGz/QG11/0J0gf9Ce4n/RIKP/1CQm/+A
sbb/ls3O/5HY3P+s8fX/s/r+/5TZ3v93vsT/Zpib/3err/+Lys//nsbF/2qptPButL+fH1Zg/Qo2
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
AADAAwAAwAMAAMADAADAAwAAwAMAAP//AAA=')
	#endregion
	$formChooseOU.MaximizeBox = $False
	$formChooseOU.MinimizeBox = $False
	$formChooseOU.Name = "formChooseOU"
	$formChooseOU.StartPosition = 'CenterParent'
	$formChooseOU.Text = "Choose Active Directory OU"
	$formChooseOU.add_Load($FormEvent_Load)
	$formChooseOU.add_Shown($formChooseOU_Shown)
	#
	# cb_AdvancedFeatures
	#
	$cb_AdvancedFeatures.Location = '19, 11'
	$cb_AdvancedFeatures.Name = "cb_AdvancedFeatures"
	$cb_AdvancedFeatures.Size = '137, 24'
	$cb_AdvancedFeatures.TabIndex = 3
	$cb_AdvancedFeatures.Text = "Advanced Features"
	$cb_AdvancedFeatures.UseVisualStyleBackColor = $True
	$cb_AdvancedFeatures.add_CheckedChanged($cb_AdvancedFeatures_CheckedChanged)
	#
	# Treeview
	#
	$Treeview.Anchor = 'Top, Bottom, Left, Right'
	$Treeview.ImageIndex = 1
	$Treeview.ImageList = $imagelist
	$Treeview.Location = '19, 37'
	$Treeview.Name = "Treeview"
	$System_Windows_Forms_TreeNode_1 = New-Object 'System.Windows.Forms.TreeNode' ("Active Directory Hierarchy", 1, 1)
	$System_Windows_Forms_TreeNode_1.ImageIndex = 1
	$System_Windows_Forms_TreeNode_1.Name = "Active Directory Hierarchy"
	$System_Windows_Forms_TreeNode_1.SelectedImageIndex = 1
	$System_Windows_Forms_TreeNode_1.Tag = "root"
	$System_Windows_Forms_TreeNode_1.Text = "Active Directory Hierarchy"
	[void]$Treeview.Nodes.Add($System_Windows_Forms_TreeNode_1)
	$Treeview.SelectedImageIndex = 1
	$Treeview.Size = '301, 442'
	$Treeview.TabIndex = 1
	$Treeview.add_AfterLabelEdit($CreateOU)
	$Treeview.add_BeforeCheck($Treeview_BeforeCheck)
	$Treeview.add_AfterCheck($Treeview_AfterCheck)
	$Treeview.add_BeforeExpand($Treeview_BeforeExpand)
	$Treeview.add_NodeMouseClick($Treeview_NodeMouseClick)
	$Treeview.add_DoubleClick($Treeview_DoubleClick)
	#
	# buttonOK
	#
	$buttonOK.Anchor = 'Bottom, Right'
	$buttonOK.DialogResult = 'OK'
	$buttonOK.Location = '245, 500'
	$buttonOK.Name = "buttonOK"
	$buttonOK.Size = '75, 23'
	$buttonOK.TabIndex = 0
	$buttonOK.Text = "OK"
	$buttonOK.UseVisualStyleBackColor = $True
	$buttonOK.add_Click($buttonOK_Click)
	#
	# imagelist
	#
	$Formatter_binaryFomatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
	#region Binary Data
	$System_IO_MemoryStream = New-Object System.IO.MemoryStream (,[byte[]][System.Convert]::FromBase64String('
AAEAAAD/////AQAAAAAAAAAMAgAAAFdTeXN0ZW0uV2luZG93cy5Gb3JtcywgVmVyc2lvbj00LjAu
MC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWI3N2E1YzU2MTkzNGUwODkFAQAA
ACZTeXN0ZW0uV2luZG93cy5Gb3Jtcy5JbWFnZUxpc3RTdHJlYW1lcgEAAAAERGF0YQcCAgAAAAkD
AAAADwMAAABwCgAAAk1TRnQBSQFMAgEBBAEAAWABAAFgAQABEAEAARABAAT/AQkBAAj/AUIBTQE2
AQQGAAE2AQQCAAEoAwABQAMAASADAAEBAQABCAYAAQgYAAGAAgABgAMAAoABAAGAAwABgAEAAYAB
AAKAAgADwAEAAcAB3AHAAQAB8AHKAaYBAAEzBQABMwEAATMBAAEzAQACMwIAAxYBAAMcAQADIgEA
AykBAANVAQADTQEAA0IBAAM5AQABgAF8Af8BAAJQAf8BAAGTAQAB1gEAAf8B7AHMAQABxgHWAe8B
AAHWAucBAAGQAakBrQIAAf8BMwMAAWYDAAGZAwABzAIAATMDAAIzAgABMwFmAgABMwGZAgABMwHM
AgABMwH/AgABZgMAAWYBMwIAAmYCAAFmAZkCAAFmAcwCAAFmAf8CAAGZAwABmQEzAgABmQFmAgAC
mQIAAZkBzAIAAZkB/wIAAcwDAAHMATMCAAHMAWYCAAHMAZkCAALMAgABzAH/AgAB/wFmAgAB/wGZ
AgAB/wHMAQABMwH/AgAB/wEAATMBAAEzAQABZgEAATMBAAGZAQABMwEAAcwBAAEzAQAB/wEAAf8B
MwIAAzMBAAIzAWYBAAIzAZkBAAIzAcwBAAIzAf8BAAEzAWYCAAEzAWYBMwEAATMCZgEAATMBZgGZ
AQABMwFmAcwBAAEzAWYB/wEAATMBmQIAATMBmQEzAQABMwGZAWYBAAEzApkBAAEzAZkBzAEAATMB
mQH/AQABMwHMAgABMwHMATMBAAEzAcwBZgEAATMBzAGZAQABMwLMAQABMwHMAf8BAAEzAf8BMwEA
ATMB/wFmAQABMwH/AZkBAAEzAf8BzAEAATMC/wEAAWYDAAFmAQABMwEAAWYBAAFmAQABZgEAAZkB
AAFmAQABzAEAAWYBAAH/AQABZgEzAgABZgIzAQABZgEzAWYBAAFmATMBmQEAAWYBMwHMAQABZgEz
Af8BAAJmAgACZgEzAQADZgEAAmYBmQEAAmYBzAEAAWYBmQIAAWYBmQEzAQABZgGZAWYBAAFmApkB
AAFmAZkBzAEAAWYBmQH/AQABZgHMAgABZgHMATMBAAFmAcwBmQEAAWYCzAEAAWYBzAH/AQABZgH/
AgABZgH/ATMBAAFmAf8BmQEAAWYB/wHMAQABzAEAAf8BAAH/AQABzAEAApkCAAGZATMBmQEAAZkB
AAGZAQABmQEAAcwBAAGZAwABmQIzAQABmQEAAWYBAAGZATMBzAEAAZkBAAH/AQABmQFmAgABmQFm
ATMBAAGZATMBZgEAAZkBZgGZAQABmQFmAcwBAAGZATMB/wEAApkBMwEAApkBZgEAA5kBAAKZAcwB
AAKZAf8BAAGZAcwCAAGZAcwBMwEAAWYBzAFmAQABmQHMAZkBAAGZAswBAAGZAcwB/wEAAZkB/wIA
AZkB/wEzAQABmQHMAWYBAAGZAf8BmQEAAZkB/wHMAQABmQL/AQABzAMAAZkBAAEzAQABzAEAAWYB
AAHMAQABmQEAAcwBAAHMAQABmQEzAgABzAIzAQABzAEzAWYBAAHMATMBmQEAAcwBMwHMAQABzAEz
Af8BAAHMAWYCAAHMAWYBMwEAAZkCZgEAAcwBZgGZAQABzAFmAcwBAAGZAWYB/wEAAcwBmQIAAcwB
mQEzAQABzAGZAWYBAAHMApkBAAHMAZkBzAEAAcwBmQH/AQACzAIAAswBMwEAAswBZgEAAswBmQEA
A8wBAALMAf8BAAHMAf8CAAHMAf8BMwEAAZkB/wFmAQABzAH/AZkBAAHMAf8BzAEAAcwC/wEAAcwB
AAEzAQAB/wEAAWYBAAH/AQABmQEAAcwBMwIAAf8CMwEAAf8BMwFmAQAB/wEzAZkBAAH/ATMBzAEA
Af8BMwH/AQAB/wFmAgAB/wFmATMBAAHMAmYBAAH/AWYBmQEAAf8BZgHMAQABzAFmAf8BAAH/AZkC
AAH/AZkBMwEAAf8BmQFmAQAB/wKZAQAB/wGZAcwBAAH/AZkB/wEAAf8BzAIAAf8BzAEzAQAB/wHM
AWYBAAH/AcwBmQEAAf8CzAEAAf8BzAH/AQAC/wEzAQABzAH/AWYBAAL/AZkBAAL/AcwBAAJmAf8B
AAFmAf8BZgEAAWYC/wEAAf8CZgEAAf8BZgH/AQAC/wFmAQABIQEAAaUBAANfAQADdwEAA4YBAAOW
AQADywEAA7IBAAPXAQAD3QEAA+MBAAPqAQAD8QEAA/gBAAHwAfsB/wEAAaQCoAEAA4ADAAH/AgAB
/wMAAv8BAAH/AwAB/wEAAf8BAAL/AgAD//8A/wD/AP8ABQAS/wz0Bv8C9Ab/A/QS/wp0BHMD/wwq
A/8D7AHrAW0B9wL/AQcC7AHrAXIBbQH0Af8KdARzAv8BdAGaA3kBegd5AXMD/wFRARwBdANzBVEB
KgP/AfcBBwGYATQBVgH3Av8BvAHvAQcBVgE5AXIB9AH/AXQBmgN5AXoHeQFzAv8BeQyaAXQD/wF0
ApkCeQN0A1IBKgP/Ae8BBwHvAngBkgLxAwcBeAFYAesB9AH/AXkCmgVLBZoBdAL/AXkMmgF0A/8B
mQIaAaAEmgJ6AXkBUgP/Ae8CBwHvAZIC7AFyAe0CBwLvAewB9AH/AXkCmgFLA1EBKgWaAXQC/wF5
AaALmgF0A/8BmQIaAaAEmgJ6AXkBUgP/AQcB7wL3Au0BeAE1AXgB7wP3AewB9AH/AXkBoAGaAXkB
mQJ5AVEFmgF0Av8BeQGgC5oBdAP/AZkCGgGgBJoCegF5AVID/wEHA+8B9wHtAZgBeAGZAQcD7wHs
Av8BeQGgAZoCmQGgAXkBUgWaAXQC/wGZAaALmgF0A/8BmQIaAaAEmgJ6AXkBUgP/AbwD8wG8AZIB
BwHvAQcB8QLzAfIB7QL/AZkBoAGaAZkBeQGaAXkBUgWaAXQC/wGZAaALmgF0A/8BmQEaAZoCmQZ5
AVID/wG8AQcC7wH3A+0B7wIHAu8B7QL/AZkBoAGaAnkBdAJSBZoBdAL/AZkBwwaaAaAEmgF0A/8B
mQEaAZkDGgOaAVIBeQFSA/8CvAIHAvcCBwO8AgcBkgL/AZkBwwGaBHQBeQGgBJoBdAL/AZkBwwOa
AqABmQWaAXQD/wGZARoBmQL2BMMBUgF5AVID/wK8AesB7AIHAvMB8AG8Ae0BbQEHAfcC/wGZAcMD
mgKgAZkFmgF0Av8BmQWgAZoCdAV5A/8BmQIaAvYEwwFYAXkBUgP/AbwBBwKSAe8B9wKSAe8BvAHv
AZIB7wH3Av8BmQWgAZoCdAV5Av8BeQGaBBoBdAOaApkBmgF5A/8BmQMaApkDeQFYAXkBUgP/A/QB
8gG8AfECvAHvAfAE9AL/AZkBmgQaAXQDmgKZAZoBeQL/AZkGeQGaAvYB1gG0AZoBeQP/AVEBHAF5
A3QBUgRRASoG/wH0AbwB9wESAewB7wHwBv8BGwZ5AZoC9gHWAbQBmgGZCP8BmgZ5AZoD/wxRBv8B
9AG8AQcC7wH3AfEM/wHDBnkBw0H/AUIBTQE+BwABPgMAASgDAAFAAwABIAMAAQEBAAEBBgABARYA
A///AAIACw=='))
	#endregion
	$imagelist.ImageStream = $Formatter_binaryFomatter.Deserialize($System_IO_MemoryStream)
	$Formatter_binaryFomatter = $null
	$System_IO_MemoryStream = $null
	$imagelist.TransparentColor = 'Transparent'
	#
	# ContextMenu
	#
	[void]$ContextMenu.Items.Add($changeDomainToolStripMenuItem)
	[void]$ContextMenu.Items.Add($newOUToolStripMenuItem)
	$ContextMenu.Name = "ContextMenu"
	$ContextMenu.Size = '170, 26'
	#
	# changeDomainToolStripMenuItem
	#
	$changeDomainToolStripMenuItem.Name = "changeDomainToolStripMenuItem"
	$changeDomainToolStripMenuItem.Size = '169, 22'
	$changeDomainToolStripMenuItem.Text = "Change Domain..."
	$changeDomainToolStripMenuItem.add_Click($changeDomainToolStripMenuItem_Click)
	#
	# newOUToolStripMenuItem
	#
	$newOUToolStripMenuItem.Name = "newOUToolStripMenuItem"
	$newOUToolStripMenuItem.Size = '203, 22'
	$newOUToolStripMenuItem.Text = "New Organizational Unit"
	$newOUToolStripMenuItem.add_Click($newOUToolStripMenuItem_Click)
	#endregion Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $formChooseOU.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$formChooseOU.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$formChooseOU.add_FormClosed($Form_Cleanup_FormClosed)
	#Show the Form
	$formChooseOU.ShowDialog() | Out-Null
    return $SelectedObject

} #End Function Choose-ADOrganizationalUnit
