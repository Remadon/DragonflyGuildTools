Keystone Profiler

This script is writtin in PowerShell and is intended for use in Microsoft Windows systems. Your Milage May Vary in other environments. 
You agree to the terms listed in the LICENSE file in the root of the DragonflyGuildTools repository, it's the unlicense.

We utiilize the ImportExcel module, you can read more about that here: https://github.com/dfinke/ImportExcel

Using the script is simple. The script will expect the presense of a CSV file named players.csv. You can find a sample of this file within the Sample Files directory of this repository. players.csv contains information the script will need to process data about the players you want to get data for, namely, the name of the character you want to get data from, and the realm that character is on. For example:

hathlo,wyrmrest-accord
daddythicc,wyrmrest-accord
vegor,zuljin

Will result in the script pulling data for those characters from those realms. The script will then create an excel spreadsheet containing a summary cover sheet, and sheets with detailed data for each player. 

You can see a sample Keystone report in the Sample Files directory, as well.

The tool utilizes RaiderIO, if you are receiving stale data or incorrect data, please ensure that character has been updated on the platform, as you may be receiving incomplete or out of data data from the API.
