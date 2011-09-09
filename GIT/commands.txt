:: GIT Commands

:: GIT: ADD AND COMMIT ALL FILES
git add -f . && git commit -a -m "Commit Message"

:: add files to staging area: 
git add -f .

:: commit and add commit note:
commit -a -m "Commit Note"

:: ignore file updates
:: http://blog.pagebakers.nl/2009/01/29/git-ignoring-changes-in-tracked-files/
git update-index --assume-unchanged <file>

:: ignore folder updates
git update-index --assume-unchanged <FOLDER>\

:: stop ignoring file updates
:: http://blog.pagebakers.nl/2009/01/29/git-ignoring-changes-in-tracked-files/
git update-index --no-assume-unchanged <file>

:: stop ignoring folder updates
git update-index --no-assume-unchanged <FOLDER>\
