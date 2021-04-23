# Description

This Form allows you top search for emails that were or were not responded to (among other things)

# How to use
### Import the form
* Download the form [LastVerbProperties.cfg](../../../raw/main/forms/LastVerbProperties/LastVerbProperties.cfg)
* Currently, importing the forms _requires_ `.ico` file(s) to be present. Either:
  * Place the form into your Office _FORMS_ folder (see [Finding you Outlook install folder](#finding-you-outlook-install-folder)) (_requires admin rights !_)
  * Modify the `LargeIcon` and `SmallIcon` properties with a valid path to a `.ico` file
    * `.ico` files are present in your Office `FORMS` folder (see [Finding you Outlook install folder](#finding-you-outlook-install-folder))
    * Other `.ico` files you provide
* Go to:
  > `Ribbon` > `File` > `Options` > `Advanced` > `Developers` > `Custom Forms` > `Manage Forms`

* Ensure the target library (right panel) is set to `Personal Forms` (or go to the right side `Set...`)
* Click `Install...` and navigate to where you deposited your `LastVerbProperties.cfg` file
* The end result should look like this:      
  ![image](https://user-images.githubusercontent.com/23620704/115908580-a532f180-a46a-11eb-8340-94d4cbbe9307.png)
  

# Exemple use: Configure a Search Folder to display emails you need to reply to
### Configure search folder
* (Optional) Enable the QueryBuilder (see: [Enable QueryBuilder](../settings/README.md#enables-query-builder-in-outlook))
* Go to the list of Outlook folders, and find `Search Folders`, then:
  > `Right click` > `New Search Folder` > `Custom` > `Create a Custom Search Folder` > `Choose`
* Name: `Need Reply`
* Browse: Select your `Inbox` and `Search subfolders` only
* Criteria:
  * Add a Categorie:
    > `More choices` > `Categories` > `New` > `Need Reply` > `OK` > `OK`
  * Use the form:
    > `Query Builder` (or `Advanced`) > `Field` > `Forms` > 
    > 
    > `Personal Forms` > `Custom` > `Properties` > `LastVerbProperties` > `Add ->` > `Close`
  * Define Criteria
    > `Field` > `LastVerbProperties` > `LastVerbExecuted` > `not equal to` > `102`
    
    Repeat last step, with values `103`, `104`, `108`, `549`
    
    ![image](https://user-images.githubusercontent.com/23620704/115911730-b978ed80-a46e-11eb-8b91-f82155b87344.png)
    
  * Validate everything
### Use search folder
* Go to some emails and categorize them as `Need Reply` : They should show up in your search folder
* Reply to some: They should disapear from your search folder

Understanding the Search Criteria : [PidTagLastVerbExecuted Property](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/87a8b6b8-59a4-4859-9dcd-8b0f36e3d729)


# Finding you Outlook install folder
One simple method would be to:
* Search Outlook in Start Menu
  > `Right Click` > `Open File Location`
  > 
  > If Shortcut: > `Properties` > `Target` > `File Location`
* Your `FORMS` folder should be at this location. You should install your file in : `FORMS\xxxx` (where `xxxx` is a 4 digit number, like `1033`, though any existing number can be used
 
