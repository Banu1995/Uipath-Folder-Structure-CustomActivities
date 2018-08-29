# Uipath-FolderStructure-CustomActivities
Folder-Structure-CustomActivities is a custom activity designed for getting the detailed structure of a folder in excel sheet. 
It will dispaly the folder structure of one or two different directories, file count based on the type of the file and a comparison sheet of two directories.
  * Install the custom activity ("Folder.Structure.Activities.0.7.1.3.nupkg") provided in the repository in UiPath.
  * Installed custom activity for comparing two directories is displayed as "Package comparer" and for one directory is displayed as           "Package comparer single" under "PackageComparerActivities" in activities panel.
  * This custom activities consists of 3 inputs, "FolderPath1", "FolderPath2", "SaveAsPath".
  * Provide two folder paths which needs to be compared, in the each respective "FolderPath1" and "FolderPath2" ("Sample UiPath Code           Snippet" is attached in repository for further reference).  
  * Provide a path with file name (For example "D:\abcxyz.xlsx") in "SaveAsPath" in which the final excel file can viewed as the output       (Sample output file is "FolderStructureDemo.xlsx" attached for reference).
  * For single directory using "Package comparer single" under "PackageComparerActivities" in activities panel will have only "FolderPath"
    and "SaveAsPath".

