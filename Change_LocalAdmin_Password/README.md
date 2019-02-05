## Description

This script will allow you to change a local user account password.

## Usage

### Running Locally

```ps
.\Change-LA-Password.ps1 -User LOCALADMIN
```

This will pass the user `LOCALADMIN` into the script. If you do not specify the `-User` switch,
you will still be prompted by the shell to enter a user.

### Updating `computers.txt`

This file is the main control center of this script.
Any computer hostnames entered there will be passed to the `ADSI` interface.

## Flow

*Step 1:*
Update `computers.txt` with the relevant hostnames.

*Step 2:*
Run the script.

## Output

*Object Type:*

```ps
$Object = New-Object -TypeName PSObject -Property @{
        ComputerName = $Computer
        Connection = $Connection
        PasswordChange = $Status
}
```

*Sample Output (Pass/Fail)*

```ps
ComputerName          Connection  PasswordChange
------------          ----------  --------------
VeryImportantServer01 ONLINE      SUCCESS
VeryImportantServer02 OFFLINE     FAILED
```
