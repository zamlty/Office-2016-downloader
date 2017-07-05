$url = 'http://officecdn.microsoft.com/pr/5440FD1F-7ECB-4221-8110-145EFAA6372F/Office/Data'

# Office 2016 Builds
# 64256AFE-F5D9-4F86-8936-8840A6A4F5BE    Deferred Channel, RTM
# 5440FD1F-7ECB-4221-8110-145EFAA6372F    Insider Channel

Invoke-WebRequest -Uri "$url/v64.cab" -OutFile 'v64.cab'
expand '.\v64.cab' -F:"VersionDescriptor.xml" .\ 2>&1>$null

[Xml]$XmlContent = Get-Content '.\VersionDescriptor.xml'
$ver = $XmlContent.Version.Available.I640Version
$des = Split-Path -Parent $MyInvocation.MyCommand.Definition
$des += "\Office_$ver`_x64_zh-cn\Office\Data"
rm '.\VersionDescriptor.xml'

echo "version = $ver"
echo 'Start downloading:'
echo "$url/v64.cab"
echo "$url/v64_$ver.cab"
echo "$url/$ver/i640.cab"
echo "$url/$ver/i642052.cab"
echo "$url/$ver/s640.cab"
echo "$url/$ver/s642052.cab"
echo "$url/$ver/stream.x64.x-none.dat"
echo "$url/$ver/stream.x64.zh-cn.dat"
pause

Invoke-WebRequest -Uri "$url/v64_$ver.cab" -OutFile "v64_$ver.cab"
md "$des\$ver" 2>&1>$null
mv *.cab -Force $des 2>&1>$null

.\aria2c -c -Z -P "$url/$ver/{i640, i642052, s640, s642052}.cab" -d "$des\$ver"
.\aria2c -c -Z -P -x16 "$url/$ver/stream.x64.{zh-cn, x-none}.dat" -d "$des\$ver"
