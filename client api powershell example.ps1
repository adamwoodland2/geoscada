# For help with the Geo SCADA Client API, in the Start menu under the Geo SCADA folder there is a "Client API" document

# Load the assembly
Add-Type -path "C:\Program Files\Schneider Electric\ClearSCADA\ClearScada.Client.dll"

# Connect to the server
$serveraddress = "127.0.0.1"
$port = 5481
$username = "adamtest"
$password = "adamtest01"

try
{
    $node = New-Object ClearScada.Client.ServerNode($clearscada.ConnectionType.Standard, $serveraddress, $port)
    $clearscada = New-Object ClearScada.Client.Simple.Connection("test")
    $clearscada.Connect($node)
    $clearscada.LogOn($username, $password)
}
catch
{
    Write-Host "Unable to connect to Geo SCADA:"
    Write-Host $_
}

# You can just get an object here you know the fullname and set a property
$object = $clearscada.GetObject("Data.AI")
if ($object -ne $null)
{
    try 
    {
        $object.SetProperty("Units", "Wibbles") # Random choice of property that exists on analogue points
    }
    catch 
    {
        Write-Host "An error occurred:"
        Write-Host $_
    }
}
else
{
    Write-Host "Object not found"
}

# Or find the users on the system (that you have permission to see) and process each one
$users = $clearscada.GetObjects("CDBUser", "")
foreach ($user in $users)
{
    if ($user.Name.ToUpper() -ne $username.ToUpper()) # Be careful changing your own user incase you make a config error!
    {
        try
        {
            # Note email address, etc are on an aggregate so must specify that (ContactConfig) too
            $user.SetProperty("AccessMask", 7) # BitMask: 1 Desktop, 2 Web, 4 PIN
            $user.SetProperty("ContactConfig.EmailAddress", "someone@somewhere.com")
            $user.SetProperty("ContactConfig.VoicemailNumber", "01234567890") # Or PagerId if a configured service
            # You can't set the PIN (or password)) directly like you can with pretty much the other properties, must use special functions
            $clearscada.ChangeUserVoicemailPin($user.id, "", (Get-Random -Minimum 0 -Maximum 9999).ToString('0000'))
            Write-Host "Updated" $user.GetProperty("FullName")
        }
        catch
        {
            Write-Host "Error updating" $user.GetProperty("FullName")
            Write-Host $_
        }
    }
    else 
    {
        Write-Host "Skipping" $user.GetProperty("FullName") "as the current user"
    }
  
}

# And if necessary create a new user and set to external authentication
$user = $clearscada.GetObject("Users.yetanothertest")
If ($user -eq $null)
{
    try
    {
        $clearscada.CreateObject("CDBUser", $clearscada.GetObject("Users").Id, "yetanothertest")
    }
    catch
    {
        Write-Host "Error creating user"
        Write-Host $_
    }
    
}
$user.SetProperty("WindowsAuth", $true)

# Usergroups can also be a pain
$usergroups = $user.GetProperty("UserGroupNames") # Read-only field
$requiredusergroup = "Users.Usergroups.Admins"
if ($requiredusergroup -in $usergroups)
{
    Write-Host "User already in usergroup, no action taken"
}
else
{
    try
    {
        # You can remove a user from a group by removing the group from the array and writing the new array back to the user object
        [string[]] $usergroups += $requiredusergroup # Make sure you use [string[]] else PowerShell converts the array to object and so errors

        $user.SetProperty("UserGroupIds", $usergroups)
        Write-Host "User added to usergroup"
    }
    catch
    {
        Write-Host "Error adding user to group"
        Write-Host $_
    }
}

# Close the connection
$clearscada.Disconnect()