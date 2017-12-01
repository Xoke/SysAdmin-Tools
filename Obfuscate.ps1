# ask for credentials
$cred = Get-Credential
$pass = $cred.Password
$user = $cred.UserName

# create random encryption key
$key = 1..32 | ForEach-Object { Get-Random -Maximum 256 }

# encrypt password with key
$passencrypted = $pass | ConvertFrom-SecureString -Key $key

# turn key and password into text representations
$secret = -join ($key | ForEach-Object { '{0:x2}' -f $_ })
$secret += $passencrypted

# create code
$code  = '$i = ''{0}'';' -f $secret 
$code += '$cred = New-Object PSCredential(''' 
$code += $user + ''', (ConvertTo-SecureString $i.SubString(64)'
$code += ' -k ($i.SubString(0,64) -split "(?<=\G[0-9a-f]{2})(?=.)" |'
$code += ' % { [Convert]::ToByte($_,16) })))'

# write new script
$editor = $psise.CurrentPowerShellTab.files.Add().Editor
$editor.InsertText($code)
$editor.SetCaretPosition(1,1) 
