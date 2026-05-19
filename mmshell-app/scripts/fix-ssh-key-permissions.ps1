# 修复SSH私钥文件权限脚本
# 使用方法: .\fix-ssh-key-permissions.ps1 -KeyPath "C:\Users\28706\.ssh\id_ed25519"

param(
    [Parameter(Mandatory=$true)]
    [string]$KeyPath
)

# 检查文件是否存在
if (-not (Test-Path $KeyPath)) {
    Write-Error "文件不存在: $KeyPath"
    exit 1
}

Write-Host "正在修复私钥文件权限: $KeyPath" -ForegroundColor Green
Write-Host ""

try {
    # 第一步：移除所有继承的权限
    Write-Host "[1/4] 移除继承权限..." -ForegroundColor Cyan
    icacls $KeyPath /inheritance:r | Out-Null
    
    # 第二步：移除所有现有的权限条目
    Write-Host "[2/4] 清除现有权限..." -ForegroundColor Cyan
    $acl = Get-Acl $KeyPath
    $acl.SetAccessRuleProtection($true, $false)
    Set-Acl $KeyPath $acl
    
    # 第三步：获取当前用户信息
    $username = $env:USERNAME
    $userDomain = $env:USERDOMAIN
    $fullUsername = "${userDomain}\${username}"
    Write-Host "[3/4] 设置用户权限: $fullUsername" -ForegroundColor Cyan
    
    # 第四步：只授予当前用户完全控制权限
    icacls $KeyPath /grant:r "${fullUsername}:(F)" | Out-Null
    
    Write-Host ""
    Write-Host "权限修复成功！" -ForegroundColor Green
    Write-Host ""
    
    # 验证权限
    Write-Host "当前权限设置：" -ForegroundColor Cyan
    icacls $KeyPath
    
    Write-Host ""
    Write-Host "现在可以正常使用该私钥文件进行SSH连接了。" -ForegroundColor Green
} catch {
    Write-Error "权限修复失败: $_"
    exit 1
}
