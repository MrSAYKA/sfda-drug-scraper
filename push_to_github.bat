@echo off
REM Push webscrapper to GitHub

echo ==================================================
echo Pushing webscrapper to GitHub...
echo ==================================================

cd /d C:\Users\imsvg\IdeaProjects\webscrapper

REM Initialize git if not already done
echo.
echo 1. Initializing git repository...
git init

REM Configure git (if needed)
echo.
echo 2. Configuring git user...
git config user.name "MrSAYKA"
git config user.email "your-email@example.com"

REM Add all files
echo.
echo 3. Adding all files...
git add .

REM Commit changes
echo.
echo 4. Committing changes...
git commit -m "SFDA Drug Scraper - Web scraping with Selenium and Excel export"

REM Add remote origin
echo.
echo 5. Adding remote repository...
git remote add origin https://github.com/MrSAYKA/webscrapper.git

REM Set main branch
echo.
echo 6. Setting main branch...
git branch -M main

REM Push to GitHub
echo.
echo 7. Pushing to GitHub...
git push -u origin main

echo.
echo ==================================================
echo Done! Your code has been pushed to GitHub.
echo ==================================================
pause

