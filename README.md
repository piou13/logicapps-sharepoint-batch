# logicapp-sharepoint-batch
This sample explore a way to use some SharePoint and Logic App features combined for background jobs.
## Introduction
Sometime, we need to run some background processes or Jobs to execute some tasks on SharePoint content. Regarding the data, we quickly face 2 main challenges : Limits and Performances.

Back in the day, we liked to use some local or cloud scripting like PowerShell, Azure WebJobs or Azure Function and we had to handle these different aspects of limits and performances.

Logic App offers a great alternative to these approaches because we can leverage some very useful features for this kind of scenario:

 - Its connectors facilitate the use of different API 
 - We can control and it takes care of the reliability (i.e: by using Retry-Policy)
 - We can control errors handling
 - We can monitor and alert
 - We 'only' focus on the control flow and messages exchanges

On the other hand,

 - it's a new environment that impose a new way of thinking (at least for me ^^) and transposing some concepts like fetching big content or batching operations could be burdensome.
 - We don't necessary have the same control on data operations and transformations compared to a pure scripting approach (i.e: the platform could be hard to extend, we can only use the set of predefined function, cannot easily manipulate the concept of local variables, etc...etc...)

This sample show a way to implement fetching over 5000 items (by overcoming the SharePoint REST limit of 5000 items max by response), applying some transformation and business logic using Liquid when Logic App OOTB actions are too limited, and batching operation using the `_api/$batch` endpoint of SharePoint.
## The Scenario
For this sample, let's create a simple scenario where we have a document library named *LogicAppSharePointBatch* with a custom string column named *FolderCode*.
![list1](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list1.PNG)
We need to get all folders from a documents library  and apply them a 'custom code' using a little 'dummy' pattern.
We want the *FolderCode* to be:

**FOLDER**`[FolderId]` **(**`[folder_path_values_sorted_alphabetically]`**)**
i.e: for a folder called "foo" with ID 4 located inside a folder "toto" at the library root, we have the following code: "FOLDER4 (foo, toto)".
I know, it's a stupid pattern, but it's just for demo purpose ^^

So, after the job, we expect any *FolderCode* to be updated with this pattern.
![enter image description here](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list2.PNG)
## The Logic App
todo
## The Liquid Template
todo
## Install the sample
todo