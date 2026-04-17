# Read the file
$content = Get-Content -Path "d:\MMShell0414\mmshell-app\src\App.tsx" -Raw

# Step 1: Remove the top progress bar
$pattern1 = '(?s)(                      </div>\r?\n                      )\{sftpProgress.*?\}(.*?\r?\n                      )<ul className="sftp-list"'
$replacement1 = '$1<ul className="sftp-list"'
$content = $content -replace $pattern1, $replacement1

# Step 2: Add progress bar below the list
$pattern2 = '(?s)(                      </ul>\r?\n                    </div>)'
$replacement2 = @'
                      </ul>
                      {sftpProgress && (
                        <div className="sftp-progress">
                          <div className="sftp-progress-info">
                            <span className="sftp-progress-file">{sftpProgress.file}</span>
                            <span className="sftp-progress-speed">{sftpProgress.speed}</span>
                          </div>
                          <div className="sftp-progress-bar">
                            <div className="sftp-progress-fill" style={{ width: `${sftpProgress.percent}%` }} />
                          </div>
                          <span className="sftp-progress-percent">{sftpProgress.percent}%</span>
                        </div>
                      )}
                    </div>
'@
$content = $content -replace $pattern2, $replacement2

# Write back
Set-Content -Path "d:\MMShell0414\mmshell-app\src\App.tsx" -Value $content -NoNewline
Write-Host "Modified successfully!"
