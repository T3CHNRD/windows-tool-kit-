Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Write-Output 'KillerTools encryption-inspired helper: AES encrypt/decrypt text and compute file/text hashes locally.'
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Encryption and Hash Helper'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(820,620)
$form.Font = New-Object System.Drawing.Font('Segoe UI',10)
$tabs = New-Object System.Windows.Forms.TabControl; $tabs.Dock='Fill'; $form.Controls.Add($tabs)
$tabText = New-Object System.Windows.Forms.TabPage; $tabText.Text='Text AES'
$tabHash = New-Object System.Windows.Forms.TabPage; $tabHash.Text='Hash'
$tabs.TabPages.AddRange(@($tabText,$tabHash))
$input = New-Object System.Windows.Forms.TextBox; $input.Multiline=$true; $input.ScrollBars='Vertical'; $input.SetBounds(12,12,760,140); $input.Text='Type text to encrypt or paste encrypted Base64 text to decrypt.'
$key = New-Object System.Windows.Forms.TextBox; $key.SetBounds(12,165,760,26); $key.Text='change-this-secret'
$output = New-Object System.Windows.Forms.TextBox; $output.Multiline=$true; $output.ScrollBars='Vertical'; $output.SetBounds(12,245,760,250); $output.ReadOnly=$true
$enc = New-Object System.Windows.Forms.Button; $enc.Text='Encrypt AES'; $enc.SetBounds(12,205,130,32)
$dec = New-Object System.Windows.Forms.Button; $dec.Text='Decrypt AES'; $dec.SetBounds(152,205,130,32)
$tabText.Controls.AddRange(@($input,$key,$enc,$dec,$output))
function Get-KeyBytes([string]$Secret) { $sha=[Security.Cryptography.SHA256]::Create(); try { $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Secret)) } finally { $sha.Dispose() } }
$enc.Add_Click({ try { $aes=[Security.Cryptography.Aes]::Create(); $aes.Key=Get-KeyBytes $key.Text; $aes.GenerateIV(); $bytes=[Text.Encoding]::UTF8.GetBytes($input.Text); $ct=$aes.CreateEncryptor().TransformFinalBlock($bytes,0,$bytes.Length); $output.Text=[Convert]::ToBase64String($aes.IV + $ct); $aes.Dispose() } catch { $output.Text=$_.Exception.Message } })
$dec.Add_Click({ try { $raw=[Convert]::FromBase64String($input.Text.Trim()); $aes=[Security.Cryptography.Aes]::Create(); $aes.Key=Get-KeyBytes $key.Text; $aes.IV=$raw[0..15]; $ct=$raw[16..($raw.Length-1)]; $pt=$aes.CreateDecryptor().TransformFinalBlock($ct,0,$ct.Length); $output.Text=[Text.Encoding]::UTF8.GetString($pt); $aes.Dispose() } catch { $output.Text="Unable to decrypt: $($_.Exception.Message)" } })
$hashInput = New-Object System.Windows.Forms.TextBox; $hashInput.Multiline=$true; $hashInput.ScrollBars='Vertical'; $hashInput.SetBounds(12,12,760,140); $hashInput.Text='Type text here, or choose a file.'
$hashOut = New-Object System.Windows.Forms.TextBox; $hashOut.Multiline=$true; $hashOut.ScrollBars='Vertical'; $hashOut.SetBounds(12,245,760,250); $hashOut.ReadOnly=$true
$hashText = New-Object System.Windows.Forms.Button; $hashText.Text='Hash Text'; $hashText.SetBounds(12,165,120,32)
$hashFile = New-Object System.Windows.Forms.Button; $hashFile.Text='Hash File'; $hashFile.SetBounds(142,165,120,32)
$tabHash.Controls.AddRange(@($hashInput,$hashText,$hashFile,$hashOut))
$hashText.Add_Click({ $bytes=[Text.Encoding]::UTF8.GetBytes($hashInput.Text); foreach($alg in 'SHA256','SHA512','MD5'){ $h=[Security.Cryptography.HashAlgorithm]::Create($alg); $hashOut.AppendText("${alg}: " + ([BitConverter]::ToString($h.ComputeHash($bytes)).Replace('-','')) + [Environment]::NewLine); $h.Dispose() } })
$hashFile.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; if($dlg.ShowDialog() -eq 'OK'){ $hashOut.Text="File: $($dlg.FileName)`r`n"; foreach($alg in 'SHA256','SHA512','MD5'){ $hashOut.AppendText("${alg}: " + (Get-FileHash -LiteralPath $dlg.FileName -Algorithm $alg).Hash + [Environment]::NewLine) } } })
[void]$form.ShowDialog()
Write-Output 'Encryption/hash helper closed.'
exit 0
