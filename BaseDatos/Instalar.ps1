param([ValidatePattern('^Clever[A-Za-z0-9_]{0,90}$')][string]$Base='CleverSoftPro6',[string]$CarpetaDatos,[switch]$DatosDemo)
$ErrorActionPreference='Stop'
$directorioSQL=Split-Path -Parent $MyInvocation.MyCommand.Path
foreach($f in @('Esquema.sql','DatosBase.sql','DatosDemo.sql')){if(!([IO.File]::ReadAllText((Join-Path $directorioSQL $f),[Text.Encoding]::UTF8)).TrimEnd().EndsWith('-- FIN_ARCHIVO_SQL')){throw "Archivo incompleto: $f"}}
function Abrir($b){$c=New-Object Data.SqlClient.SqlConnection "Server=localhost\SQLEXPRESS;Database=$b;Integrated Security=True";$c.Open();return $c}
function Ejecutar($cn,$s,$tx=$null){$c=$cn.CreateCommand();$c.CommandText=$s;$c.CommandTimeout=600;if($tx){$c.Transaction=$tx};try{[void]$c.ExecuteNonQuery()}finally{$c.Dispose()}}
$cn=Abrir 'master'
try{
 $c=$cn.CreateCommand();$c.CommandText='SELECT COUNT(*)FROM sys.databases WHERE name=@b';[void]$c.Parameters.AddWithValue('@b',$Base);if($c.ExecuteScalar()){throw 'La base ya existe; no se sobrescribe.'}
 $archivos=''
 if($CarpetaDatos){$d=[IO.Path]::GetFullPath($CarpetaDatos);[void][IO.Directory]::CreateDirectory($d);$mdf=Join-Path $d ($Base+'.mdf');$ldf=Join-Path $d ($Base+'_log.ldf');if((Test-Path -LiteralPath $mdf)-or(Test-Path -LiteralPath $ldf)){throw 'Los archivos de destino ya existen.'};$archivos=" ON PRIMARY(NAME=N'$Base',FILENAME=N'$($mdf.Replace("'","''"))',SIZE=16MB,FILEGROWTH=64MB) LOG ON(NAME=N'$($Base)_log',FILENAME=N'$($ldf.Replace("'","''"))',SIZE=16MB,FILEGROWTH=64MB)"}
 Ejecutar $cn "CREATE DATABASE [$Base]$archivos COLLATE Modern_Spanish_CI_AS; ALTER DATABASE [$Base] SET COMPATIBILITY_LEVEL=100; ALTER DATABASE [$Base] SET AUTO_CLOSE OFF; ALTER DATABASE [$Base] SET AUTO_SHRINK OFF; ALTER DATABASE [$Base] SET RECOVERY SIMPLE; ALTER DATABASE [$Base] SET PAGE_VERIFY CHECKSUM;"
}finally{$cn.Dispose()}
$cn=Abrir $Base;$tx=$cn.BeginTransaction()
try{
 foreach($f in @('Esquema.sql',$(if($DatosDemo){'DatosDemo.sql'}else{'DatosBase.sql'}))){$s=[IO.File]::ReadAllText((Join-Path $directorioSQL $f),[Text.Encoding]::UTF8);foreach($lote in [regex]::Split($s,'(?im)^\s*GO\s*$')){if($lote.Trim()){Ejecutar $cn $lote $tx}}}
 Ejecutar $cn 'DBCC CHECKCONSTRAINTS WITH ALL_CONSTRAINTS' $tx
 $tx.Commit();$tx=$null;"INSTALACION_CORRECTA: $Base"
}catch{if($tx){$tx.Rollback()};throw}finally{$cn.Dispose()}
