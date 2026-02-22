# Your one-stop Altium shop!
Welcome to the MIL-Altium repo. This repo contains plenty of resources for MIL Altium projects.

# Initial Setup
<details>
<summary>## Setting up Altium scripts</summary>

This repo uses custom scripts to make certain things in Altium easier. In order for the scripts to operate properly, please follow these steps.
1) Clone this repo to **C:/MIL-Altium**. This must be where the repo exists on your device.
2) Open Altium.
3) Go to the Preferences menu (gear icon in the top-right of Altium).
4) Click "Load" near the bottom of the Preferences menu. Select C:/MIL-Altium/Configuration/MIL_Config.DXPPrf. Click "Apply All" in the smal window that appears. It will take a second to update all Altium config settings. Allow "Altium Settings Elevator" to run with administrative priviledges when prompted.
5) Restart Altium to ensure all changes have taken effect.
6) Click on the MIL dropdown menu on your toolbar. This should be visible from the schematic or PCB editor.
7) Click the "Init MIL Scripts" command. This will configure the scripts to work on your device. Follow the instructions that appear.
8) All done! See the MIL Altium Script Guide in the Guides folder to learn about available scripts and how they work.
</details>


<details>
<summary>## Installing SamacSys Library Loader</summary>

SamacSys Library Loader is a very useful tool for importing component footprints. Installation of this tool is not required, but highly recommended, especially if you need to import a component into Altium in the future. Once installed, you can import pretty much any component on Mouser.com super quickly (assuming it has a part library available).
1) Navigate to MIL-Altium/Scripts/Altium Library Loader 2.2 Setup.exe
2) Run this executable and follow the instructions it gives.
3) You will need to create a SamacSys/Component Search Engine account if you do not have one already.
4) Once installation has completed, restart Altium.
5) All done! See the SamacSys Guide in the Guides folder to learn more about how to use this program within Altium.
</details>


# What's in this repo?
## Part Libraries
Parts used in your projects should exist within the part libraries in this repo. If they don't exist yet, add them!
[add guide for how to add new components to library, required steps, pull request, etc.]

## Design Rules & Stackups
Your PCB project must use the design rules provided in this folder. Additionally, make sure to select the proper stackup for your project (2-layer or 4-layer).

## Scripting
meow

## Templates
meow