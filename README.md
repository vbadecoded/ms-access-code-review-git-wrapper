<H1 align="center">A VBA Decoded Project</H1>
<p align="center">
  <img src="https://www.vbadecoded.com/images/logo.svg" width="200" />
</p>

---

# Microsoft Access Code Review / GIT Wrapper
#### An Access DB management tool that allows you to quickly and safely develop your DB while tracking changing via GIT.

# INSTRUCTIONS
*This is the full process to contribute.*
*Please read all instructions*
 
| # | First Time Setup Steps |
| ----------- | ----------- |
| 1 | Install and Set Up GIT | 
| 2 | Setting up Database Repository |
| 3 | Drive Setup |
| 4 | Gain GitHub Access | 
| 5 | Clone Repositories |

### 1. Install and set up [GIT](https://git-scm.com/install/windows)
- Download/Install
- Log in using GitHub
### 2. Setting up Database Repository
- Create a new folder for your git repository
- Inside the folder, open git bash
- Run `git init` to initialize the repository
- Add the master database file (the .accdb) to the folder
- Run `git add .` to add the file to git
- Run `git commit -m "initial commit"`
### 3. Drive Setup (network drive based production repos only)
- IF your production MS Access Database is on a shared drive on your network, you'll need to make sure that drive is mapped.
- In this repository, you'll want to add a file in the root called .productionLocation
- Map Prod Location of the database
- For example, I use something like this: "\\\data\test\AccessDB\build\" to Drive Letter: A
### 4. Gain [GitHub](https://github.com/) Access (if necessary)
- Log in/Sign Up
- Send a Code Owner your username to gain access to repository as a contributor
### 5. Clone Repositories
- First, clone Code Review : https://github.com/vbadecoded/ms-access-code-review-git-wrapper
- Then, clone the repository you want to work on
- *I typically name the folder to match the repository name*
 
---
| # | Steps to Publish New Changes |
| ----------- | ----------- |
| 1 | GIT Process |
| 2 | Accept/Reject the Changes | 
| 3 | Release Changes to Production Repository |
 
### GIT Process
1. Open Code Review Database
2. Select Repository to work on
3. Click "Status" to check status of repository changes
   This is the same as the Bash comand:
	```bash
	git status
	```
 *Before you start working, make sure you are on the correct GIT branch*
4. Click "Enable Shift" to allow using Shift Bypass on Database
5. Shift + Click "Open Database" to bypass startup procedures
6. Do your work on the MS Access Database
7. Click Clean + Decompose Database
8. Click Status to see what files were changed
*NOTE: there may be many unexpected changes listed, especially in forms. This is expected, since even just opening a form in the dev version will alter binary attributes of the forms files. Typically you can ignore this. I usually focus on reviewing the .mod files behind the forms*
 
#### Accept/Reject the Changes
1. Review changes using the Git GUI program or the Diff button
   This is the same as the Bash comands:
	```bash
	git diff
	```
2. Commit and then Push (if you have a remote repo) to your branch (I do not recommend working on Master/Main branch)
   This is the same as the Bash comand:
	```bash
	git commit
	git push origin branch-name
	```
 
#### Release Changes to Production Repository
*You can merge the changes however you like. But remember, the files you are tracking changes on are not perfectly representative of the composed database*
1. AFTER PUSH, switch to master branch
   This is the same as the Bash comand:
	```bash
	git checkout master
	```
2. Merge your branch
   This is the same as the Bash comand:
	```bash
	git merge branch-name
	```
3. "Push" to make changes public in GitHub
      This is the same as the Bash comand:
	```bash
	git push origin master
	```
4. Publish (Pull) down in production location to actually change the production front end
   This is the same as the Bash comand (in production repository):
	```bash
	git pull origin master
	```
 
# CODE OWNER - How to accept/reject changes
1. Have a local version of this repository
2. Move to branch that is submitted for review
3. Review changes
4. Accept or Reject changes
5. Open a Pull request
	- Recompose accepted changed into .accdb file and clean DB if necessary
	- Revert other changes
6. Send feedback to contributor
7. Merge pull request with accepted changes in Main branch
