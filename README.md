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

On the other hand, it's a new environment that impose a new way of thinking (at least for me ^^) and transposing some concepts like fetching big content or batching operations could be burdensome.

This sample show a way to implement fetching over 5000 items (by overcoming the SharePoint REST limit of 5000 items max by response) and batching operation using the `_api/$batch` endpoint.
## The Scenario
