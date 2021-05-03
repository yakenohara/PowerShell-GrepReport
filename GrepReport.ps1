# <License>------------------------------------------------------------

#  Copyright (c) 2021 Shinnosuke Yakenohara

#  This program is free software: you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation, either version 3 of the License, or
#  (at your option) any later version.

#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.

#  You should have received a copy of the GNU General Public License
#  along with this program.  If not, see <http://www.gnu.org/licenses/>.

# -----------------------------------------------------------</License>

#
# <CAUTION>
# この .ps1 スクリプトファイル自体のテキストエンコードは、
# UTF-8 (※BOM有り※) としておかないと、日本語文字列の検索がヒットしない。
# </CAUTION>
#

# <Settings>----------------------------------------------------------------

# note
# 同じディレクトリを意味する `.\` を付けると、
# `Get-ChildItem` コマンドレットの `-Exclude` に反映されない
$outPathWithoutLineNo = 'Grep Report without line no.md'
$outPathWithLineNo = 'Grep Report with line no.md'

$incFiles = @(
    '*.md'
    '*.svg'
)
$excFiles = @(
    $PSCommandPath,        # この .ps1 スクリプトファイルは除外する
    $outPathWithoutLineNo, # 出力済みファイルは除外する
    $outPathWithLineNo,    # 出力済みファイルは除外する
    'synonyms.md',
    'images.md'
)
$regexNGwords = @(
    'ギリシア(?!神話)',
    'ホメイニ(?!ー)',
    '(?<!マグレブ諸国の)アルカイダ'
)
$enc_name = "utf-8" # 出力ファイルのテキストエンコード
$indent = '    '
# ---------------------------------------------------------------</Settings>

# 検索対象となる `System.IO.FileInfo` オブジェクトリストを作成
$fInfoToGrep = 
    Get-ChildItem -Path .\ -Recurse -File -Include $incFiles -Exclude $excFiles | # 指定ファイル名、指定除外ファイル名で `System.IO.FileInfo` オブジェクトリストを取得
    Sort-Object -Property FullName # フルパスの名称で sort

# 対象件数が 0 だった場合は終了
if ($fInfoToGrep -eq $null){ # 対象件数が 0 だった場合
    Write-Host 'File to search not found.'
    return
}

# 出力先ファイル StreamWriter を開く
try{
    $enc_obj = [Text.Encoding]::GetEncoding($enc_name)
    
    if ($enc_obj.CodePage -eq 65001){ # for utf-8 encoding with no BOM
        $outFileWithoutLineNo = New-Object System.IO.StreamWriter($outPathWithoutLineNo, $false)
        $outFileWithLineNo  = New-Object System.IO.StreamWriter($outPathWithLineNo, $false)
        
    } else {
        $outFileWithoutLineNo = New-Object System.IO.StreamWriter($outPathWithoutLineNo, $false, $enc_obj)
        $outFileWithLineNo  = New-Object System.IO.StreamWriter($outPathWithLineNo, $false, $enc_obj)
    }
    
} catch { # 出力先ファイル StreamWriter を開けなかった場合
    Write-Error ("[error] " + $_.Exception.Message)
    try{
        $outFileWithoutLineNo.Close()
        $outFileWithLineNo.Close()
    } catch {}
    return
}

# 検索処理
for ($idx = 0 ; $idx -lt $regexNGwords.count ; $idx++){
    
    Write-Host "Processing $($idx + 1) of $($regexNGwords.count):``$($regexNGwords[$idx])``" #todo 0埋め
    $strNGword = "# Searched result about ``$($regexNGwords[$idx])``"
    $outFileWithoutLineNo.WriteLine($strNGword)
    $outFileWithLineNo.WriteLine($strNGword)

    # 文字列として取得
    # note `Microsoft.PowerShell.Commands.MatchInfo` のオブジェクトリストとして取得される
    $mchInfoPerLine = Select-String -Pattern $regexNGwords[$idx] -Path $fInfoToGrep -AllMatches

    if ( $mchInfoPerLine.count -eq 0 ){ # ヒットしなかった場合
        $str_tmp = 'Nothing matched.'
        Write-Host $str_tmp
        $outFileWithoutLineNo.WriteLine($str_tmp)
        $outFileWithLineNo.WriteLine($str_tmp)

    }else{ # 1件以上ヒットが存在する場合
        
        $amountOfMatches = 0 # 検索ヒット数を初期化
        
        # 検索ヒット数のカウント
        for ($idxOfLines = 0 ; $idxOfLines -lt $mchInfoPerLine.count ; $idxOfLines++){
            $amountOfMatches += $mchInfoPerLine[$idxOfLines].Matches.count
        }
        $str_tmp = "$amountOfMatches matched."
        Write-Host $str_tmp
        $outFileWithoutLineNo.WriteLine($str_tmp)
        $outFileWithLineNo.WriteLine($str_tmp)

        # 検索結果のファイル毎表示
        $fileOfInterest = ''
        for ($idxOfLines = 0 ; $idxOfLines -lt $mchInfoPerLine.count ; $idxOfLines++){

            # 参照しているファイルが前回のループから変わっているかどうか確認
            $relPathOfHittedFile = $mchInfoPerLine[$idxOfLines].RelativePath($PSScriptRoot) # この .ps1 スクリプトファイルからみた相対パスの取得
            if ($fileOfInterest -ne $relPathOfHittedFile){ # 参照しているファイルが前回のループから変わっている場合
                $fileOfInterest = "$relPathOfHittedFile"
                Write-Host $fileOfInterest
                $strForFileName = "## $fileOfInterest"
                $outFileWithoutLineNo.WriteLine($strForFileName)
                $outFileWithLineNo.WriteLine($strForFileName)
            }

            $mchInfoPerLine[$idxOfLines].Matches | ForEach-Object {
                $strForWithLineNo = " - Line:$($mchInfoPerLine[$idxOfLines].LineNumber), Index:$($_.Index)"
                Write-Host $strForWithLineNo
                $outFileWithLineNo.WriteLine($strForWithLineNo)
            }

            $strForWithoutLineNo = $mchInfoPerLine[$idxOfLines].Line
            Write-Host $strForWithoutLineNo
            $outFileWithoutLineNo.WriteLine('```')
            $outFileWithoutLineNo.WriteLine($strForWithoutLineNo)
            $outFileWithoutLineNo.WriteLine('```')
            
            $outFileWithLineNo.WriteLine('```')
            $outFileWithLineNo.WriteLine($strForWithoutLineNo)
            $outFileWithLineNo.WriteLine('```')

        }
    }
}

# file close
$outFileWithoutLineNo.Close()
$outFileWithLineNo.Close()
