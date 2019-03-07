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

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list1.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list1.PNG)

We need to get all folders from a documents library  and apply them a 'custom code' using a little 'dummy' pattern.
We want the *FolderCode* to be:

**FOLDER**`[FolderId]` **(**`[folder_path_values_sorted_alphabetically]`**)**
i.e: for a folder called "foo" with ID 4 located inside a folder "toto" at the library root, we have the following code: "FOLDER4 (foo, toto)".
I know, it's a stupid pattern, but it's just for demo purpose ^^

So, after the job, we expect any *FolderCode* to be updated with this pattern.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list2.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list2.PNG)

## The SharePoint List
For our test, we inject more than 5000 folders to make sure we hit the list threshold, but, more important, to make sure the REST request for all folders does not content 'all folders' but the 5000 first ones.
These folders doesn't have to be at the root. We can have a folder structure.
Here's my situation:

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list3.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list3.PNG)

## The Logic App
The process is divided into four main steps:

 1. Fetching the Documents Library Folder's items to get all of them.
 2. Checking Items for transformation logic in order to prepare to batch update items.
 3. If items that need to be updated are detected, then batch-process these items.
 4. Generate and send the batch report to a recipient (could be something else, like leveraging Azure Alert).

> I won't go through all the Logic App structure. I just highlight the
> important point I had to tackle.

A good starting point to deal with different aspect of SharePoint REST API is here: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)

**Error Handling**
I like to use scopes to organize the process sequences because everything is explained here ;) : [https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-control-flow-run-steps-group-scopes](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-control-flow-run-steps-group-scopes)

My Logic App overall structure with error handling looks like this:

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list4.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list4.PNG)

As you can see, each logical decomposition represented by a scope has a parallel branch to manage error that occurred somewhere in the scope. Up to you to plug any custom logic to manage your errors.

**Logic App Variables**
To help you to understand the data manipulation inside the flow, here is the explanation for every variable we use in the Logic App:

 - *SiteAbsoluteUrl*: 
 Stores the site collection absolution URL. You don't really need this one because you can pass this information using a better way, but once again, it's for demo purpose.
 - *ListName*:
 Stores the name of the SharePoint Documents Library. Same consideration as SiteAbsoluteUrl.
 - *ItemsArray*:
 Stores the fetched items from the list.
 - *NextLink*:
Stores if the response has a next page of results.
 - *NextLinkUrl*:
Stores the URL of the next results page if exists.
 - *BatchRequestBody*:
Stores the content of the batch request body for the SharePoint batch query.
 - *ChangeSetRequestBody*:
 Stores the content of the changeset request body for the SharePoint batch query.
 - *BatchBoundary*:
 Stores the batch boundary identifier.
 - *ChangeSetBoundary*:
 Stores the changeset boundary identifier.
 - *BatchResponse*:
 Stores the SharePoint batch response.
 - *FoldersToProcess*:
 Stores the list of folders that need to be processed.

**Step 1: Fetch items**

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list5.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list5.PNG)

The challenge here occurs when we have more than 5000 items in the documents library. The SharePoint REST API can returns up to 5000 items max if we specify the `$top=5000` parameter, but no more. In this case, SharePoint returns pages of response containing 5000 items. Here, the next result page can be accessed by getting the `odata.nextLink` value from the response.

So, we start with a REST query to get the information we need, something like:

    _api/web/lists/getbytitle('[sp_documents_library]')/items?$top=5000&$Expand=FieldValuesAsText&$Filter=startswith(ContentTypeId, '0x0120')&$Select=Id,FolderCode,FieldValuesAsText/FileRef

The interesting point here is inside the *IfPaged* condition. We implement a recursion while we have next page  to get all the 5000 items from the page and merge them into the *ItemsArray* array. So, once the recursion done, the ItemsArray will contain every items from the list and not only the 5000 first ones. In the next steps, we will use this array as 'datasource' to proceed.

But the transformation step, we need to convert this ItemsArray to a JSON object because we need to provide to Liquid a JSON object he is able to understand.

The only tricky point I had here was to replace any attribute using a dot (.) in their names (i.e: metadata like `odataid`, `odata.etag`, ...) because it's like Liquid doesn't manage attribute with dots. So I take the original JSON provided by *ItemArray* and I replace any "odata." string by "odata". This is why we have an additional *FixedItemsJson* action.

**Step 2: Check items**

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list6.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list6.PNG)

In this step, we want to check all the items to filter and transform according to our rules.
First, we want to keep only those that don't have a *FolderCode* defined.
Then, we want to transform some information to generate the *FolderCode*.
Finally, we want to send back the result in a way that can be easily used by the following steps in Logic App.

Here, to get to the goal, Logic App's built-in actions like 'Select' are too limited for our need, and it's very hard to implement any custom logic using them. Also, we can use intensively Logic App's built-in loops like For-each or Until, but they are slow and very challenging when they are used in conjunction with variables. For the good and the bad about that, see: [https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-control-flow-loops](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-control-flow-loops) and [https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-create-variables-store-values](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-create-variables-store-values)

This is why I decided to investigate on Liquid because I heard it was a kind of transformation engine for JSON like XSLT could be for XSL. Sounds interesting!

More information about Liquid later but my objective here is to start from the JSON containing my items (*FixedItemsJson*), only keep the ones with no FolderCode defined and make transformations to return the following JSON object as output:

    {
      "FolderCode": "[item_folder_(wonderful)_code_pattern]",
      "FolderEtag": [item_folder_etag],
      "FolderId": [item_id],
      "FolderName": "[folder_name]"
    }

Let's call this object a "FolderToProcessObject".
In order to do that, I defined and stored a Liquid template to manage these operations. Then, I consume it from the Liquid action "Transform JSON to JSON".

For each item sent back by Liquid, I create the related "FolderToProcessObject" object and append it to my *FolderToProcess* array variable.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list7.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list7.PNG)

**Step 3: Batch update SharePoint items**
For sure, this step is only needed if we have folders to proceed.
So first, we check the length of the *FolderToProcess* array variable to make sure it has content.

The next actions before the batch request are made to dynamically build the REST request. I won't go over how to make batch request to SharePoint, good information can be find here: [https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis), but the challenge here was to format correctly the HTTP request sent to the SharePoint $batch endpoint...To make it short:

The good news:
We don't care about authentication! No need to manage an access token or other X-RequestDigest stuff... All is managed by the SharePoint connector. :).

The tricky:
It took me a while before handling the way Logic App manages newlines and the HTTP request requirements... I mean, in Logic App  .. 1 line + 2 lines ...

i.e: Just to illustrate my point on one action,

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list8.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list8.PNG)

Between the end of ChangeSet header and HTTP request, we need two lines (in fact hit 'ENTER' 3 times..) and between the HTTP request header and the body, we need only one line. Anyway, if you face the error `"The header value must be of the format header name: header value"`, dig this point ^^.

Finally, we build the complete batch request body using a Compose action called FinalRequestBody to pass to the SharePoint batch request.

To actually execute the batch request, we use the so useful SharePoint action "Send HTTP Request to SharePoint". Be aware of the Content-Type header declaration that uses the multipart and the batch boundary.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list9.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list9.PNG)

The final step of this sequence is to receive and format the batch response and append it to our *BatchResponse* array variable.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list10.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list10.PNG)

**Step 4: Report**
Entirely optional, this step gather information from the *BatchResponse* array and send them to a mail recipient using the SMTP connector.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list11.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list11.PNG)

## The Liquid Connector
**Prerequisites**
To use the Liquid connector, we need a basic Azure Integration Account to consume from our Logic App. More info here: [https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-create-integration-account](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-create-integration-account)

Also,  we need a bit of understanding about Liquid's syntax and the way it works. 2 resources are useful for that: from Microsoft ([A great help is here: [https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-liquid-transform](https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-liquid-transform)
](A%20great%20help%20is%20here:%20%5Bhttps://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-liquid-transform%5D%28https://docs.microsoft.com/en-us/azure/logic-apps/logic-apps-enterprise-integration-liquid-transform%29)) and from the Liquid official editor ([https://shopify.github.io/liquid/](https://shopify.github.io/liquid/)).

**The Liquid Template**
If you have read the Liquid documentation on Azure, you know that your Liquid template needs to be stored in an Integration Account service. Once, uploaded and your Logic App connected to this Integration Account, you can consume it from the Logic App. So, let's go directly to the point and look at the template code.

![https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list12.PNG](https://github.com/piou13/logicapp-sharepoint-batch/blob/master/docs/list12.PNG)

`content` : it's the representation of our provided JSON.

`{% for c in content %}` : Parse all our items.

`{% unless  c.FolderCode %}` : Keep only item with no FolderCode.

`{% assign  fileref = c.FieldValuesAsText.FileRef  | Split: '/' %}` : Create an array to get all the path to the item.

`{% capture  paths %}{% for  path in fileref  offset: 3 %}{{ path }}|{% endfor %}{% endcapture %}` and `{% assign  filerefSorted = paths  | Split: '|'  | Sort: | Join: ', ' %}` : Get rid of the three first index (because from the *FileRef* path, [0] represent "sites" (the managed path), the [1] [site_name] and the [2] [library_name] and return a sorted array with the remaining values we convert to a string joined by `', '`.

`"FolderCode": "FOLDER{{ c.Id }} ({{ filerefSorted }})"` : We generate our FolderCode.
## Install the sample
todo