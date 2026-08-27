> [!NOTE]
> Altium scripts are not yet totally functional. If something isn't working, let Russell know.

# Your one-stop Altium shop!
Welcome to the MIL-Altium repo. This repo contains plenty of resources for MIL Altium projects.


## Initial Setup
#### 1) Clone the repo to your root directory
In order to function properly, this repo must be cloned to C:/MIL-Altium. If you're not familiar with Git or how to clone a repo, you can also run the following script.

Press "Win+R", type "cmd", then copy/paste the following into your terminal.

`cd /d "%USERPROFILE%\Downloads" && curl -fsSL -o clone-mil-altium.bat "https://raw.githubusercontent.com/uf-mil-electrical/MIL-Altium/main/Scripts/Bash%20Scripts/clone-mil-altium.bat" && powershell -Command "Unblock-File -Path clone-mil-altium.bat" && clone-mil-altium.bat`


#### 2) Installing SamacSys Library Loader
SamacSys Library Loader is a very useful tool for importing component footprints. Installation of this tool is not required, but highly recommended, especially if you need to import a component into Altium in the future. Once installed, you can import pretty much any component on Mouser.com super quickly (assuming it has a part library available).
<details>
<summary>Steps</summary>
  
1) Navigate to MIL-Altium/Scripts/Altium Library Loader 2.2 Setup.exe
2) Run this executable and follow the instructions it gives.
3) You will need to create a SamacSys/Component Search Engine account if you do not have one already.
4) Once installation has completed, restart Altium.
5) All done! See the SamacSys Guide in the Guides folder to learn more about how to use this program within Altium.
</details>


#### 3) Setting up Altium scripts
This repo uses custom scripts to make certain things in Altium easier. In order for the scripts to operate properly, please follow these steps.
<details>
<summary>Steps</summary>
  
1) Open Altium.
2) Go to the Preferences menu (gear icon in the top-right of Altium).
3) Click "Load" near the bottom of the Preferences menu. Select C:/MIL-Altium/Configuration/MIL_Config.DXPPrf. Click "Apply All" in the smal window that appears. It will take a second to update all Altium config settings. Allow "Altium Settings Elevator" to run with administrative priviledges when prompted.
4) Restart Altium to ensure all changes have taken effect.
5) Open a schematic document. Go to File >> New >> Schematic.
6) Click on the MIL dropdown menu on your toolbar.
7) Click the "Init MIL Scripts" command. This will configure the scripts to work on your device. Follow the instructions that appear.
</details>
