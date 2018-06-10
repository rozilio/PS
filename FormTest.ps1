# Header Information
#############################################################################################
#
#   Script: Sample.ps1
#
#   Author: Gareth J. Edwards
#
#     Date: 01/07/2015
#
#    About: This script is a sample script
#
#    Usage: This script is used as a template
#
# Requirements:
#Requires -Version 2
#
# Versions:
# 01/07/2015 - 1.0 - Initial release creation by Gareth J Edwards
#
#############################################################################################

#region User Entered Variables

$FormTitle = "Sample Form"
$VersionNo = "1.0"

$ColWidthOdd = 100
$ColWidthEven = 150
$VertSpace = 6
$HorizSpace = 16
$RowHeight = 20

$ButtonWidth = 70
$ButtonHeight = 20

# ComboBox lists
$cboSampleText3_List = @("List Sample 1","List Sample 2","List Sample 3","List Sample 4")
$cboSampleText4_List = @("List Sample 5","List Sample 6","List Sample 7","List Sample 8")

#region Customer Logo
$CustLogo = "/9j/4AAQSkZJRgABAQEAYABgAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/7AARRHVja3kAAQAEAAAAZAAA
/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwEC
AgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAHgBFAwEi
AAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNR
YQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6
g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/E
AB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRC
kaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJ
ipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMR
AD8A/fwnFNWTNeEfGn9unTfAfjG88J+D/C/iT4m+MNNH+nWOhxD7JpJIyFvLxv3VudvO05YAjIGVzzGiftJ/HXxrua3+H/wm8MoE
aUpq3j/7ZMiKNzMy20BGAOvzDGK82Wa4fndOF5Nb8sZSS8m0mr+V7ntU+H8XKkq0+WEXquecItruoykpNdnazPpySTD8V5v42/a1
8A+CPFEmgPrM2seIoTtk0nQbC41i+hOOkkVqkhi+sm0cjJGa8t8S+KPjN8QdCm0m7uPgakd6ViaKx8V6lbzT5P8AqxJEiyLuOB8h
BPToSDc8ED4s/CzwX5XhrwH+z/pXh+33Nt0zxDc2NpGVYh2O2y2ZBByTzkHPNZ1cfUbtThJebhJ/grfn8jXD5TTSbrTjJ9EqkF97
d/uS+Z6Jb/HXxFqlsJrL4T/ECS3k+49zPpVmzj18uS8Ei/R1U+1Y/iT9qnVvA8Xna58HvitbW3UzWNtYasPf5LS7lkH4oPbNchqf
x2+O+nSmNvB/wREiorgN4/nVsMMqQDZDqCCDkZBFY9/8XP2mdSsBdWum/s+6NZyyGJLifxFd3aq45K7lVFLY5wK56mOko+77S/8A
17f6xX5nXh8pi5fvPY8v/X5J/hJ6/wDbr9DrfCv/AAUx+C/iLV/7OuPFz+HdTQ4ks/EGl3ekyQn0YzxIgPtur2vw34u0nxppMeoa
Pqen6tYTcx3NlcJcQuD6OhIP518daj+w14w/aE1S88WfEbxd8Eo55dq3V3o3ga01KRlRQo33l4zbdq4GNp4A5rc8Df8ABP69+AF0
viDwP8VvDPhV7xVBvJPAmkeVco3IXfF5JKsPRskdDXPg8bm3NfEUU4dLWjK3nFyav80dmaZXw8oKODxLjOyunzThfS9pRpxdt7e6
z68VxRXK/CO91K58MeTrHijQPFmqW74mvNIsvscOD90GLz5tp4PO/n0FFfSRbau1b+vmfF1IqMmk7+avZ/fZ/ej4X/4Lgfs/654n
s/g/c2vgfxf42+Bei+Ib7UfiV4P8Dq8epauZUVra5eKBkknjWbzWl2sG+cnIYiSP8+vg58I9a8OxfHLQfBfwT1LT9G8V/BvxJBNq
R+GGt6PqDXIh3Q2Fuby+vSVc7ThCrSlVBB8sV/QzXFfHbwh4k8b+CvsPhfWP7G1L7TFK0nnvb+bCp+ePzUVmjzx8wUnjHQmlOapw
bS26GlGDq1FGT36vU/ED/glN+zlp3gv9rr4O6hqHwj8Uw65YXkLTyyfCDV4Pstz5DBppL+81U2cYjYmQz/ZAfkzFGkhTbi/C/wDZ
S/aW/Z5/4I7axfeF9H8SeKPh/wDFnQ9S0rxr8PdR0u4j1fwpefaZYIdWsbfZ5rIyRwtKgXkMWZWXEsP7ef8ACv8A4kN8Q7fUf+Ev
006DHoB0iTTfsrBpLwpuOoeZ2kEwVQmCvllj97Aqx8HPhn408C/CK80XxF4wbxJ4hlEgttVaMqYS0KKvBJb5ZQ7ck8N2GFGEcZJz
ceR9e3S3n1udE8HGMb+0V9NLPrfy6dfXS+p+D/7Tf7LOqXn7WmqXHiz4XeLb6xTwP4OtLaWT4caz4gjW4h8P2Ec8Y+x3ll5boylG
3PJhkK7VKnPuXjP9nTWvjx4O/Zy+Evhz9mXSfFWl2nhDUtStLjxrb+IvBWmaTcG+uDPut7O8mSB5liRka5d7mUMrOw31+oei/BL4
qWWjeFYbrx4JJNK15b/U0+1Sv9qsRHGpt/NMYaTLpJJhgo/e7M4UE7S/D74pRXnjC4bxZo90urapaX+iWhikgj0qGG4XzLRnAYtH
NbRoGbbkSyTEDayhc446T3g/w7X7/L/gG0stSelRfdLul2+fot7n4y+FP2NPHHwf1b9qvw5dfC3UPhXdf8KmudHsPCHhO31rXNA8
aXDypN9ut765eUSzRx7VEKlXxu2opWcN00n7OUegeLPhVqv7TnwR+NnxZ+Ep+C/hLS/BWn+F7S+uIfCd/DpVot/b3lrbzQSQ3DTJ
Kfn6hgCG2fuf1u8TfCj4ka/4r8WXlr4xh0mz1zRI7LS7eGWR00e8xF5kqrtG4ZWUhtwb5wMDqO7+EnhrWPCHgaz03XtYbX9Us2lW
S/ZNrToZXaLIOTlYyinJJJXlmPzHaniZTny8rW+unR27313OethY06anzpvTTXqr9VbTZ677aHgf/BKXSPhXoX7PmpQfB/4P+N/g
74T/ALZmd7PxXpsthfapclE8y4AnmlnkQDbGHkIA2bV4XAK+ohRXQcZ//9k="
#endregion Customer Logo

#region Company Logo
$CompanyLogo = "/9j/4AAQSkZJRgABAQEAYABgAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/7AARRHVja3kAAQAEAAAAZAAA
/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwEC
AgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAHgBqAwEi
AAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNR
YQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6
g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/E
AB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRC
kaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJ
ipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMR
AD8A/fDUtQg0jTprq6mitrW2jaWaWVgscSKMszE8AAAkk9hX5i/td/8ABY3xV4w8S32kfCu4j8O+HLd2jTWGtlk1DUQOPMQSArDG
eqgqXIwSV5UfU3/BXbx/eeA/2IvEEdnI8MmvXdrpMsinaVhkk3SjPoyIyn2Y1+UNh4Ts/DtjDfeJWuIlmQS2ulQMI729U4IZ2IP2
eFgciRlLODlEYfMPybxC4kxlCtHL8HLk93mlLbdtJX6ba21d0kf0V4L8EZXi8LPOs0gqr5nCnBq6bSTbcftPVJX0Vm33XpHgX9tf
49XOvtcaP8QvF1/cWw86YXFwk9rEncyiYGJEPTLbRzgEZr7S/ZM/4K26X4uzovxQ1LwtoeqafavPLq1qZls9Q24Hlou0jzu+1GZX
/g/u1+bfxN+JOn+E/BdpqfjbVrbwb4TkzNpWjWFv5lxqZGR5lpZFw9wf4Wu532/LtaclQleW6D+1j8RLi50vxl4DttD+CPgPR9Qj
vIPE2u3e7UNcMMqu0KXZi8+5B2bZLbS4FUq5WbzBhq4eCcHn86ixFOo/Yvd1G2pf4Y3v87q/4Hr+KWYcHU6UsHXox+srZUVGMof4
525fWKjJr8T+gzwn+0LL8Sis3hfwZ4tvtMfldU1K1Gj2rj/ZW5K3DD3WEqfWq/jP9sTwX4F8eXnhRpNe8QeLNKtYrvU9J8N6Fe65
PpUcoJjNwbWJ1hLgEoshV3UZVSKh+D3hH4leJo9N17xp8Q9MuI5lju00nwtoC6dp7xum4RyyXTXFzJjIO6NrcnAygGQfmf8A4JU/
Ee2+GHxT/aM+FfjTULPQPi9N8Tta8ULFq2IZvEOmXjKbK/twzI1zbrHGExGcRqiodnAr9io05xX7yV36WXy6/e2fzNiKlKb/AHUO
Verb+b0X3Jeh9XfC39rLwD8YfAsviTR9eW30u31KfR5zrNrPos9veQHE0Dw3iRSq6dwV9xkc11mhfEPQvFFtFNpusaTqEM85tY5L
W8jmWSYIZDGpUkFwgLbRztBOMV+XH7Yv7R+q/tR/8Ex/iP4g8TW/gCxm0P472fhXz9PsHhs75tM8QW1j9uuTJO+/zY4YyRuGyNQm
9gAw9d/bA+BcfxI+Hfw78BR+IvBvgfVfEHxqQDWfhqv9nPpt5H4fvriyuJELt/patDaM6FvnjCDgMGrU5z7xvfGWk6drUOm3Gpaf
b6lcrvhtJblEnlXnlUJ3EcHkDsazLL4z+EdT1KOztfFHhu4vJgzRwRapA8rhVZ2IUNkgKrMfQKT0Br4e+Anxt8ffEL/gop8Jfh38
ZvDtrpfxc+Fnh/xJNeaxZ25GkeMNOnTT4odVsHKgL5jDE0HBikBHQgDK/wCCYWhaXJ8YviFrMK/CH7Bp/wAePF0Mb3lkF8RW1w8U
sSCzmEmAzec6lCmfKebGd2QAfoB4e+Jfh3xdeyW2k67ouqXEaeY0VpfRTyKn94qrE496saZ410nWtWubCz1TTru+s8+fbQXSSTQY
ODvQHcuDxyOtfkB+yhqHh7wX8RPg14n8Y6To/g7wRovxt8fpYfEDTZ1aaTVJtSvra20bUGESfY7S5MzkMZZUke0gjZY94ZfS/wBo
KbxxbeOf+ClUnwn+2L8QBpfhT7E2kg/2iIzpKfaDBs/eed9nMxQr82/bt+bFMD9Ibf41+D7zxefD8Pirw1LryytCdNTVIGvBIv3k
8oNv3DuMZFdPXy3+zVbfs1+Nf2VfhLfeF4fh23hSzm0d/DJh+zi4s9UV4vIQMv7xb0THbIpxIWMqydXFfUlID5a/4KL+D9Y+IfwK
8TXlzqFnoum+G9kug2NxAJm1zVBIFjZ0wxbJJjt4lB3SuHYMAi1+L3xn/aSt/hf40k8OeH7MfEb4vahetaNAIzqVjo98z7fKZBu/
tDUFkyDBhreNhtkM7boo/wCgH4vw6P4B0vV/iPr8Nzq0fgbS7nUrO1RVP2NY4WaV4lY4M7qGXexGFIUbcuX+PPiz/wAETLP9qTRv
GPjq+8Xr8O/i78UpRdapqXhbTon0+ysXj2f2agkVJpFdNhuLkPDLdSKzERxMLZfmsRwxg8Xj1jsXHmcVZJ6p+bW2myX33b0+6wHH
mZZdlEspy+fIpu8pJWkk0lyxe6va8ne7eislr+JvjrxFZ/Dzxre6542vrX4rfFa7bzLmK6uft2iaDIMgLdSKdt/cx44toj9ji+QO
9x+8to737JfwL8a/8FMP21/CfhG8vNQ1681i7ik1m+lb5dK0eF0Ny6quEhijiOyONAiB3iRQu4Cv0E8J/wDBpfrUevKuufHLS10W
MjjTfCbrdSLjt5l0UQ9ujDv7V+l/7B3/AATk+GH/AATv+Hlzofw+026+2aoY21fW9RkWfVNZZN2wzSKqqFXe22ONUjXcxCgsSfp7
pKyPhm23dnulpbJaQrHGgjjjUKiAYCqBgAfSvkL9qbUo/jB8TvFWn2en2vxCsPh4YP7Ts1+HNh4gXw7M9ukxhEt1dRNNMYmWYxW6
O6rLGCNzKD9hV4HZfsw+Nvhb8WfiFr/w38YeH9N0z4mX8esanp2u6JJfNpmprbRWr3VrJHPFlJIoIS0MqsA6Fg6higkRxfxN+FSf
CzwzpPhLX/G/w/XTXWe903w7bfCmO/wkKjzp4rKCR2CRiUBpAm1fNUEguM0tE+E2g6b8EdH8baX8RvgvH8P5JINY0zVrP4d2P9nC
S5aGKO6idJ9oZ28lfMGD8qgn5Rj1v4n/ALO+uap+0loHxS8I67pen69pfh688LXllq9g93Z3tlPcQXKuhjkjeKaOaBecsjI7qVB2
uvJz/wDBPfSZv2K9L+B8uprqnh6XUlvfEM17a4OrpLqTajerGkbAQeZPI+wAsIl2j5tuSAX9P8MeMvFnxM1DS4PjB4P1DxZ4Pt4j
eRjwTE93pMV4CyAn7TlBMIM4B+YRgntXmPwe8F+GfGnxlurHwf8AED4Ux+PNDW8kcRfCuGx1GFYZmsrto3eRHYJLIYZChIUyhWxv
APof7In7D2ofsrfE/WPEl140m8ZXnivwzpeleIb/AFCxEeo6vqNhPetFevIrbcfZ7sW4TbkJbRfMcYrof2fP2U/+FGXPjDWppdE1
Xxh4g1jWL+x1T7HIgsba+vGvFtCGkY7FmYFyhTzPLQkAquADiPAXwJn8YaH4s8L+G/H3wvvNL03U20nxHpVl8OrT7Kl4sUUrQ3EQ
n2M4jlibDA8MKzPAmjW/h/VYdQ8K/FT4bWuoaxr0/g43GjfDmETPqVv5vnWMrQzbkeMW8jESYAWPd0wa679lX9ivVP2X/ifqniKP
xe2uR+MdGtx4tgurVt2oa5FLNI2pW7b8QrL9omV4mDkhIMOPLO7nvh1/wTbXwH8ZPC3jq38WS2Gvaf4k1TXPE1vY2YTT/F8VzJqU
lqLiNmJS6tTqARblW3NHE0bAqyCIA5zUPh9oPg74x61eSeOvhfY/EDw9daaL26j+E8P9qxT6nI8NkyyLJ5khnkSVA8ZYZjcEja2P
sOOZo41V/nZRhmHG4+uK8o+I37MEfjf9rn4b/FFLrT7dvAul6rps9u1qWuNQ+2CERN5oYBfIKS7Qyt/x8yYK7jn1vZ7mgD//2Q=="
#endregion Company Logo

#endregion User Entered Variables

#region Script File Variables

$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir = split-path -parent $ScriptPath
$ScriptName = (Get-ChildItem $ScriptPath).Basename

#endregion Script File Variables

#region Calculated Variables

#	Different OS versions use different sizes of windows, these corrections allow the form to be correctly sized in any OS.
#	Switch -wildcard doesn't work in Powershell 2 so OS version has extra sections removed.
$OSVer = ((Get-WmiObject Win32_OperatingSystem).Version).split(".")[0..1] -join "."
Switch ($OSVer)
{
    6.3 {$FormHorizPadding = 21; $FormVertPadding = 43}		#Windows 8.1 & Windows Server 2012 R2
    6.2 {$FormHorizPadding = 21; $FormVertPadding = 43}		#Windows 8 & Windows Server 2012
    6.1 {$FormHorizPadding = 10; $FormVertPadding = 29} 	#Windows 7 & Windows Server 2008 R2
    6.0 {$FormHorizPadding = 10; $FormVertPadding = 29} 	#Windows Vista & Windows Server 2008
    5.2 {$FormHorizPadding = 21; $FormVertPadding = 43} 	#Windows Server 2003
    5.1 {$FormHorizPadding = 21; $FormVertPadding = 43} 	#Windows XP
}

$FullColWidth = $ColWidthOdd + $ColWidthEven + ($HorizSpace * 2)
$ButtonSize = "$ButtonWidth,$ButtonHeight"

$Script:TabIndex = 0
$Script:RowCount = 0
$Script:RowHigh = 0
$Script:ColCount = 0

$CompulsoryVar = @()
$EditedFieldList = @()

#endregion Calculated Variables

Function OnApplicationLoad {
# Load modules (AD etc) check for required files
return $true
}

Function ImportAssemblies {
[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[Void][reflection.assembly]::LoadWithPartialName("System.Web")
[System.Windows.Forms.Application]::EnableVisualStyles()
}

#region Object Functions

Function ObjectSize ($CellSize) {
$CellRowHeight = (($RowHeight + $VertSpace) * $RowHigh) - $VertSpace
If([bool]!($Col%2)){$ColWidth = $ColWidthEven}
Else{$ColWidth = $ColWidthOdd}
Switch($CellSize)
{
    Full {$CellWidth = $ColWidth}
    HalfLeft {$CellWidth = $ColWidth/2 - $HorizSpace/4}
    HalfRight {$CellWidth = $ColWidth/2 - $HorizSpace/4}
    Double {$CellWidth = $ColWidthOdd + $ColWidthEven + $HorizSpace}
}
$ObjectSize = "$CellWidth,$CellRowHeight"
Return $ObjectSize
}

Function ObjectLocation ($CellSize) {
If(![bool]!($Col%2)){$LocationX = $HorizSpace + ($FullColWidth * ((($Col + 1)/2)-1))}
Else{
    Switch ($CellSize){
        Full {$LocationX = ($HorizSpace * 2 + $ColWidthOdd) + ($FullColWidth * (($Col/2)-1))}
        HalfLeft {$LocationX = ($HorizSpace * 2 + $ColWidthOdd) + ($FullColWidth * (($Col/2)-1))}
        HalfRight {$LocationX = ($FullColWidth * ($Col/2)) - ($ColWidthEven/2 - $HorizSpace/4)}
        Double {$LocationX = $HorizSpace + ($FullColWidth * (($Col/2)-1))}
    }
}
$LocationY = $VertSpace + (($VertSpace + $RowHeight) * ($Row -1)) + $HeaderHeight
$ObjectLocation = "$LocationX,$LocationY"
Return $ObjectLocation
}

Function New-FormTitle ($varName) {
$object = New-Object System.Windows.Forms.Label
$object.Font = New-Object System.Drawing.Font("Arial", "16",[System.Drawing.FontStyle]::Bold)
$object.Name = $varName
$object.TextAlign = 'MiddleLeft'
$object.Text = $FormTitle
$object.AutoSize = $True

New-Variable $varName -Value $object -Scope Script
(Get-Variable formMain).Value.Controls.Add((get-variable $varName).value)

If([bool](($object.height)%2)){$objectHeight = $object.Height + 1}
Else {$objectHeight = $object.Height}
If($HeaderHeight - $objectHeight -lt 2){$Script:HeaderHeight = $objectHeight + 2}
$LocationX = $HorizSpace
$LocationY = $HeaderHeight/2 - ($objectHeight)/2
$Location = "$LocationX,$LocationY"
$object.Location = $Location
}

Function New-FormObject ($varName,$CellSize,$Text,$Style) {
$Type = $varName.substring(0,3)
Switch ($Type){
    Lbl{$TypeLong = "Label"}
    Txt{$TypeLong = "TextBox"}
    Cbo{$TypeLong = "ComboBox"}
    Chk{$TypeLong = "CheckBox"}
}

$object = New-Object System.Windows.Forms.$TypeLong
$ObjectSize = ObjectSize $CellSize
    $Object.Size = new-object System.Drawing.Size($ObjectSize)

$ObjectLocation = ObjectLocation $CellSize
    $object.Location = $ObjectLocation
$object.Name = $varName
Switch ($Type){
    Lbl {
        $object.Text = $Text
        $object.TextAlign = 'MiddleLeft'
        $object.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    Txt {
        $object.TextAlign = 'Left'
        If($Text -eq "Hidden"){$Object.Visible = $False}
        $object.TabIndex = $TabIndex; $Script:TabIndex +=1
        $Script:EditedFieldList += $object

    }
    Cbo {
        $object.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $object.TabIndex = $TabIndex; $Script:TabIndex +=1
        $object.Sorted = $True
        $Script:EditedFieldList += $object
    }
    Chk {
        $object.TabIndex = $TabIndex; $Script:TabIndex +=1
        $Script:EditedFieldList += $object
    }
}
Switch($Style){
    OutputLabel {
        $object.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $object.TextAlign = 'TopLeft'
    }
    DropDown {
        $object.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
        $object.AutoCompleteCustomSource.Add("System.Windows.Forms");
        $object.AutoCompleteCustomSource.AddRange(("System.Data", "Microsoft"));
        $object.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
        $object.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems;
    }
}

If($Text -eq "Compulsory"){$Script:CompulsoryVar += $object}

# Add new variable to form
New-Variable $varName -Value $object -Scope Script
(Get-Variable formMain).Value.Controls.Add((get-variable $varName).value)
If($Script:TabIndex -eq 1){$Script:FirstField = (get-variable $varName).name}

If($varName -like "cbo*"){
    If(Get-Variable ($VarName + "_List") -ErrorAction SilentlyContinue){
        ForEach($Entry in (Get-Variable ($VarName + "_List")).Value){((Get-Variable $varName).value).Items.Add($Entry)}
    }
}
$NewRowCount = $Row + $RowHigh - 1
If($RowCount -lt $NewRowCount){$Script:RowCount = $NewRowCount}
If($ColCount -lt $Col){$Script:ColCount = $Col}
}

Function New-FormButton ($varName,$Text,$Loc) {
$object = New-Object System.Windows.Forms.Button
If($FormSection -eq "Footer"){
    $LocationX = $FormWidthInternal - ($ButtonWidth + $HorizSpace) -($ButtonWidth + ($HorizSpace / 2)) * ($Loc - 1)
    $LocationY = $FormHeightInternal - $FooterHeight/2 - $ButtonHeight/2
}
Else{
    $ButtonWidth = $ColWidthEven/2 - $HorizSpace/4
    $ButtonSize = "$ButtonWidth,$ButtonHeight"
    $LocationX = ($FullColWidth * ($Col/2)) - ($ColWidthEven/2 - $HorizSpace/4)
    $LocationY = $VertSpace + (($VertSpace + $RowHeight) * ($Row - 1)) + $HeaderHeight
}
$Location = "$LocationX,$LocationY"
$object.Location = New-Object System.Drawing.Size($location)
$object.Size = New-Object System.Drawing.Size($ButtonSize)
$object.Text = $Text
$object.TabIndex = $Script:TabIndex; $Script:TabIndex += 1
New-Variable $varName -Value $object -Scope Script
(Get-Variable formMain).Value.Controls.Add((get-variable $varName).value)
#	((get-variable $varName).value).add_Click({$varName})
}

Function New-FormPicture ($varName,$FormSection) {
$object = New-Object System.Windows.Forms.PictureBox
$object.Image = ([System.Drawing.Image]([System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream(($$ = [System.Convert]::FromBase64String($Picture)),0,$$.Length)))))
$object.Width = $object.Image.Width
$object.Height = $object.Image.Height

If($FormSection -eq "Header"){
    $ImageLocY = ($HeaderHeight / 2) - ($object.Height / 2)
    $ImageLocX = $FormWidthInternal - $object.Width - $FormHorizPadding
}
ElseIf($FormSection -eq "Footer"){
    $ImageLocY = ($FormHeightInternal - ($FooterHeight) / 2) - ($object.Height / 2)
    $ImageLocX = $HorizSpace
}
$Location = "$ImageLocX,$ImageLocY"
$object.Location = New-Object System.Drawing.Size($Location)

New-Variable $varName -Value $object -Scope Script
(Get-Variable formMain).Value.Controls.Add((get-variable $varName).value)
}

#endregion Object Functions

#region Button Functions

Function btnCancel {
Write-Host "Cancel Button pressed"
$formMain.Close()
}

Function btnReset {
Write-Host "Reset Button pressed"
ResetForm
}

Function btnSubmit {
ForEach ($Field in $EditedFieldList){$ErrorProvider.Clear()}
$ErrorCount = 0

Write-Host "Submit Button pressed"
Start-Sleep 1

ForEach ($Field in $CompulsoryVar){
    If(!($Field.Text)){
        $ErrorProvider.SetError($Field, "No data entered, this field is compulsory");
    }
}
}

#endregion Button Functions

#region Script Functions

Function ResetForm {
ForEach ($Field in $Script:EditedFieldList) {
    $FieldType = $Field.GetType().Name
    Switch($FieldType) {
        ComboBox {$Field.SelectedIndex = -1}
        TextBox {$Field.Text = ""}
        CheckBox {$Field.Checked = $False}
    }
}
((get-variable $FirstField).value).focus()
}

#endregion Script Functions

Function LoadForm{

#region Form Objects
$formMain = New-Object System.Windows.Forms.Form
$formMain.Text = $FormTitle + " " + $VersionNo
$formMain.Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$formMain.KeyPreview = $True
$formMain.MinimizeBox = $False
$formMain.MaximizeBox = $False
$formMain.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

$ErrorProvider = New-Object System.Windows.Forms.ErrorProvider
$ErrorProvider.BlinkStyle = "NeverBlink"
$myPen = new-object System.Drawing.Pen Black

$formMain.Add_KeyDown({If($_.KeyCode -eq "Escape"){btnCancel}})
$formMain.Add_KeyDown({If($_.KeyCode -eq "Enter"){btnSubmit}})


$CustImage = ([System.Drawing.Image]([System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream(($$ = [System.Convert]::FromBase64String($CustLogo)),0,$$.Length)))))
If($HeaderHeight - $CustImage.Height -lt 10){$Script:HeaderHeight = $CustImage.Height + 10}

$DCImage = ([System.Drawing.Image]([System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream(($$ = [System.Convert]::FromBase64String($CompanyLogo)),0,$$.Length)))))
If($FooterHeight - $DCImage.Height -lt 10){$Script:FooterHeight = $DCImage.Height + 10}

New-FormTitle "lblTitle"

#endregion Form Objects

#region Form Column 1
$Col = 1
$Row = 1; $RowHigh=1
New-FormObject "lblSampleLabel1" "Full" "Label 1"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "lblSampleLabel2" "Full" "Label 2"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "lblSampleLabel3" "Full" "Label 3"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "lblSampleLabel4" "Full" "Label 4"

#endregion Form Column 1

#region Form Column 2
$Col = 2
$Row = 1; $RowHigh=1
New-FormObject "txtSampleText1" "Full" "Compulsory"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "txtSampleText2l" "HalfLeft"
New-FormObject "txtSampleText2r" "HalfRight"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "cboSampleText3" "Full"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "cboSampleText4" "Full"

#endregion Form Column 2

#region Form Column 3
$Col = 3
$Row = 1; $RowHigh=1
New-FormObject "lblSampleLabel5" "Full" "Label 5"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "lblSampleLabel6" "Full" "Label 6"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "lblSampleLabel7" "Full" "Label 7"
$Row = $Row + $RowHigh; $RowHigh=3
New-FormObject "lblSampleLabel8" "Double" "Information to the user" "OutputLabel"

#endregion Form Column 3

#region Form Column 4
$Col = 4
$Row = 1; $RowHigh=1
New-FormObject "txtSampleText5" "Full"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "txtSampleText6" "HalfLeft"
New-FormButton "btnTest" "Test"
$Row = $Row + $RowHigh; $RowHigh=1
New-FormObject "cboSampleText7" "Full"

#endregion Form Column 4

#region Form Size Calculations
$FormWidthInternal = ($ColWidthEven * ($ColCount/2)) + ($ColWidthOdd * ($ColCount/2)) + ($HorizSpace * ($ColCount + 1))
$FormWidth = $FormWidthInternal + $FormHorizPadding
$FormHeightInternal = ($RowHeight * $RowCount) + ($VertSpace * ($RowCount + 1))  + $HeaderHeight + $FooterHeight
$FormHeight = $FormHeightInternal + $FormVertPadding
$FooterStart = $FormHeightInternal - $FooterHeight
$FormMainSize = "$FormWidth,$FormHeight"
$formMain.Size = New-Object System.Drawing.Size($FormMainSize)
$formGraphics = $formMain.createGraphics()

#endregion Form Size Calculations

#region FormFooter
$FormSection = "Footer"
$ButtonNo = 1
New-FormButton "btnSubmit" "Submit" $ButtonNo
$btnSubmit.add_Click({btnSubmit})
$ButtonNo ++
New-FormButton "btnReset" "Reset" $ButtonNo
$btnReset.add_Click({btnReset})
$ButtonNo ++
New-FormButton "btnCancel" "Cancel" $ButtonNo
$btnCancel.add_Click({btnCancel})

#endregion FormFooter

#region Form Graphics
$formMain.add_paint({$formGraphics.DrawLine($mypen, 5, $HeaderHeight, $FormWidthInternal - 5, $HeaderHeight)})
$formMain.add_paint({$formGraphics.DrawLine($mypen, 5, $FooterStart, $FormWidthInternal - 5, $FooterStart)})
$Picture = $CompanyLogo
New-FormPicture "picCompany" "Footer"
$Picture = $CustLogo
New-FormPicture "picCustomer" "Header"
#endregion Form Graphics

# Show form on screen
$formMain.Add_Shown({$formMain.Activate()})
$formMain.ShowDialog()
}

#Call OnApplicationLoad to initialize
If((OnApplicationLoad) -eq $true)
{
ImportAssemblies
LoadForm | Out-Null
}