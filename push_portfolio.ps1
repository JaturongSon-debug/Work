# Portfolio Push Script
$token = Read-Host -Prompt "Please enter your GitHub Personal Access Token"
$repoUrl = "https://$token@github.com/JaturongSon-debug/Work.git"

git branch -m main
git remote set-url origin $repoUrl
git add .
git commit -m "Final professional portfolio commit"
git push -u origin main

Write-Host "Portfolio pushed successfully!" -ForegroundColor Green
