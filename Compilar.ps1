$ErrorActionPreference='Stop'
$raiz=Split-Path -Parent $MyInvocation.MyCommand.Path
$salida=Join-Path $raiz 'bin';[void][IO.Directory]::CreateDirectory($salida)
$enc=[Text.Encoding]::GetEncoding(1252)
$v=[IO.File]::ReadAllText("$raiz\Clever.vbp",$enc)
$v=[regex]::Replace($v,'(?m)^Reference=(.+)$',{param($m)$p=$m.Groups[1].Value.Trim()-split '#';if($p.Count-ge5-and![IO.Path]::IsPathRooted($p[3])){$p[3]=[IO.Path]::GetFullPath((Join-Path $raiz $p[3]))};'Reference='+($p-join'#')})
$v=[regex]::Replace($v,'(?m)^(Form|UserControl|Designer)=(.+)$',{param($m)$m.Groups[1].Value+'='+(Join-Path $raiz $m.Groups[2].Value.Trim())})
$v=[regex]::Replace($v,'(?m)^(Class|Module)=([^;]+); (.+)$',{param($m)$m.Groups[1].Value+'='+$m.Groups[2].Value+'; '+(Join-Path $raiz $m.Groups[3].Value.Trim())})
$v=$v.Replace('Path32="bin"',('Path32="'+$salida+'"'))
[IO.File]::WriteAllText("$salida\Clever-build.vbp",($v-replace "`r?`n","`r`n").TrimEnd()+"`r`n",$enc)
$log=Join-Path $salida 'compilacion.log';$inicio=Get-Date
[IO.File]::WriteAllText($log,'',$enc)
$p=Start-Process 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE' -ArgumentList @('/make',('"'+$salida+'\Clever-build.vbp"'),'/out',('"'+$log+'"')) -WindowStyle Hidden -PassThru
if(!$p.WaitForExit(20000)){throw 'Compilación en curso'}
$t=[IO.File]::ReadAllText($log,$enc);$exe=Join-Path $salida 'Clever.exe'
if($t-notmatch'ha tenido éxito'-or$t-match'Error en la generación'-or!(Test-Path $exe)-or(Get-Item $exe).LastWriteTime-lt$inicio){throw $t}
'COMPILACION_CORRECTA'
