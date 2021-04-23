# Description
Contains different settings that can be enabled by following the instructions detailled here

### Summary
* [Settings](#settings)
  * [Enables Developer Ribbon in Outlook](#enables-developer-ribbon-in-outlook)
  * [Enables Query Builder in Outlook](#enables-query-builder-in-outlook)
  * [Enables Unsafe Rules](#enables-unsafe-rules)
  * [Set Macro Security Level to Notification Only](#set-macro-security-level-to-notification-only)
* [How to use](#how-to-use)
* [Finding Outlook Version](#finding-outlook-version)

# Settings

### Enables Developer Ribbon in Outlook

![image](https://user-images.githubusercontent.com/23620704/115897908-4024cf00-a45d-11eb-9a05-b3f78a8c8006.png)
* 16.0 ([Download](../../../raw/main/settings/16.0/EnableDeveloperTools.reg))

### Enables Query Builder in Outlook
![image](https://user-images.githubusercontent.com/23620704/115900553-5bdda480-a460-11eb-968d-1eeede28d46b.png)
* 16.0 ([Download](../../../raw/main/settings/16.0/EnableQueryBuilder.reg))

### Enables Unsafe Rules
![image](https://user-images.githubusercontent.com/23620704/115901019-f76f1500-a460-11eb-8dec-7b2c1bc2dc55.png)
* 16.0 ([Download](../../../raw/main/settings/16.0/EnableUnsafeRules.reg))

### Set Macro Security Level to Notification Only
![image](https://user-images.githubusercontent.com/23620704/115903461-e4aa0f80-a463-11eb-89db-c2782e0eb79a.png)
* 16.0 ([Download](../../../raw/main/settings/16.0/SetMacroSecurityLevelNotification.reg))



# How to use
* Download the file corresponding to your version of Office/Outlook. See [Finding Outlook Version](#finding-outlook-version)
* Execute it:
  * Either by double-cliking it (requires admin rights !)
  * Or by using the command line:
    ```bat
    reg import Setting.reg
    ```

# Finding Outlook Version
See Outlook [Documentation](https://support.microsoft.com/en-us/office/what-version-of-outlook-do-i-have-b3a9568c-edb5-42b9-9825-d48d82b2257c)

Exemple for version 16.0 :

![image](https://user-images.githubusercontent.com/23620704/115899150-b37b1080-a45e-11eb-9b1a-f9c26b3d6936.png) 
