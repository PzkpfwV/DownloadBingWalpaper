param(
    $Directory = $PSScriptRoot
)
$Params = @{
    format = "js"; # 返回JSON格式结果
    idx    = 0; # 相对日期 从当前日期dd开始向前idx天(dd-idx)
    n      = 8; # 获取(dd-idx-n,dd-idx]几日内的Wallpaper
    mkt    = "zh-CN" ; # 语言/区域?
}
function ConvertTo-URLString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [HashTable[]]$HashTable
    )
    process {
        $Res = $null;
        # $Pairs = @();
        foreach ($item in $HashTable) {
            foreach ($entry in $item.GetEnumerator()) {
                $str = "{0}={1}" -f $entry.Key, $entry.Value;
                # $Pairs += "{0}={1}" -f $entry.Key, $entry.Value;
                $Res = @($Res, $str ) | Join-String -Separator "&";
            }
        }
        return $Res;
        # return ( $Pairs | Join-String -Separator "&");
    }
}

function FileNameFormat {
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [String]$filename,
        [int32]$TargetLen = "OHR.GoldenEnchantments.jpg      ".Length - 1
    )
    $Len = $filename.Length;
    if ($Len -le $TargetLen) {
        return $filename + (" " * ($TargetLen - $Len) );
    }
    else {
        $null = $filename -match "(?<name>.+)(?<Ext>\.[\S]+)";
        $PostfixStr = "~$($Matches.Ext)";
        $PostfixStrLen = $PostfixStr.Length;
        $NewName = "{0}{1}" -f ($Matches.name.Substring(0, ($TargetLen - $PostfixStrLen))), ($PostfixStr);
        return $NewName;
    }
}

$Website = "https://www.bing.com";
$URI = "$($Website)/HPImageArchive.aspx?$($Params | ConvertTo-URLString)"; # Bing 获取Wallpaper API
$Res = Invoke-WebRequest -Uri $URI;
if (-Not $Res.BaseResponse.IsSuccessStatusCode) {
    Write-Error "Network error!";
    exit 1;
}
$ImagesInfo = ($Res.Content | ConvertFrom-Json).Images;
Write-Host "Downloading Bing Wallpapers ..." -ForegroundColor:Blue;
Write-Host "Date     `t FileName                     `tStatus";
$LenFileName = "FileName                     `tStatus".Length - " `tStatus".Length; 
foreach ($item in $ImagesInfo) {
    $info = $item;
    $filename = "$($item.fullstartdate).jpg";
    $null = ($item.url -Match "id=(?<Name>[\S^&]+?.jpg)");
    $info | Add-Member `
        -MemberType NoteProperty `
        -Name "OgrinalName" `
        -Value ($Matches.Name) `
        -Force; 
    $info | Add-Member `
        -MemberType NoteProperty `
        -Name "NewName" `
        -Value ($filename) `
        -Force; 
    $info.url = -join ($Website, $info.url);
    $info.urlbase = -join ($Website, $info.urlbase);
    $info.quiz = -join ($Website, $info.quiz);
    $OutputName = FileNameFormat `
        -filename ($info.OgrinalName -replace "_ROW[\S]*.jpg", ".jpg") `
        -TargetLen $LenFileName;
    Write-Host ($item.startdate)"`t"$OutputName"`t" `
        -NoNewline;
    # $DownloadURL = "$($Website)$($item.url)";
    if ([System.IO.File]::Exists((Join-Path $Directory $filename))) {
        Write-Host "Skipped" -ForegroundColor:Green;
        continue;
    }
    try {
        Invoke-WebRequest -Uri $info.url -OutFile (Join-Path $Directory $filename);
        $info | `
            ConvertTo-Json | `
            Out-File (Join-Path $Directory "$($item.fullstartdate).json") -Encoding utf8;
        Write-Host "Success" -ForegroundColor:Green;
    }
    catch {
        Write-Host "Failed" -ForegroundColor:Green;
    }
}
Write-Host "Operations completed successfullly." -ForegroundColor:Blue;
# -Headers $Headers;
