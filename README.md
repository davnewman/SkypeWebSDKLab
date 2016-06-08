<a name="HOLTop" />
# Integrating Conversations with the Skype Web SDK #

---

<a name="Overview" />
## Overview ##

Modern-day conversations come with an expectation of being contextual and real-time. Office 365 offers a number of powerful tools that developers can use to deliver high-caliber conversations. This module will concentrate on conversation-driven development with Skype for Business. In this lab, you will convert an existing web application to integrate Skype presence and instant messaging with the Skype Web SDK. You will also see how to integrate additional conversation modalities like voice and video.


<a name="Objectives" />
### Objectives ###
In this module, you'll see how to:


- Skype-enable and existing web application using the Skype Web SDK.

<a name="Prerequisites"></a>
### Prerequisites ###

The following is required to complete this module:

- [Visual Studio Community 2015][1] or greater
- [Microsoft Office 2016][2]

[1]: https://www.visualstudio.com/products/visual-studio-community-vs
[2]: https://portal.office.com

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.

<a name="Setup" />
### Setup ###
In order to run the exercises in this module, you'll need to set up your environment first.

1. Open Windows Explorer and browse to the module's **Source** folder.
2. Right-click **Setup.cmd** and select **Run as administrator** to launch the setup process that will configure your environment and install the Visual Studio code snippets for this module.
3. If the User Account Control dialog box is shown, confirm the action to proceed.

> **Note:** Make sure you've checked all the dependencies for this module before running the setup.

<a name="CodeSnippets" />
### Using the Code Snippets ###

Throughout the module document, you'll be instructed to insert code blocks. For your convenience, most of this code is provided as Visual Studio Code Snippets, which you can access from within Visual Studio 2015 to avoid having to add it manually.

>**Note**: Each exercise is accompanied by a starting solution located in the **Begin** folder of the exercise that allows you to follow each exercise independently of the others. Please be aware that the code snippets that are added during an exercise are missing from these starting solutions and may not work until you've completed the exercise. Inside the source code for an exercise, you'll also find an **End** folder containing a Visual Studio solution with the code that results from completing the steps in the corresponding exercise. You can use these solutions as guidance if you need additional help as you work through this module.

---

<a name="Exercises" />
## Exercises ##
This module includes the following exercises:

1. [Developing with the Skype Web SDK](#Exercise2)

Estimated time to complete this module: **60 minutes**

>**Note:** When you first start Visual Studio, you must select one of the predefined settings collections. Each predefined collection is designed to match a particular development style and determines window layouts, editor behavior, IntelliSense code snippets, and dialog box options. The procedures in this module describe the actions necessary to accomplish a given task in Visual Studio when using the **General Development Settings** collection. If you choose a different settings collection for your development environment, there may be differences in the steps that you should take into account.

<a name="Exercise2"></a>
### Exercise: Developing with the Skype Web SDK ###

Skype is one of the most popular communication platforms in the world. Many organization look to **Skype for Business** to deliver their real-time communication needs. Skype for Business offers powerful SDKs to integrate real-time conversations into both web and mobile application.

In this exercise, you will convert and existing web application to integrate Skype **presence** and **instant messaging** with the **Skype Web SDK**. You will also see how to integrate additional conversation modalities like **voice** and **video**.

<a name="Ex2Task1"></a>
#### Task 1 - Setup ####

This exercise uses a starter solution that follows a help desk scenario.  It displays help desk tickets assigned to the signed in user. Your job in the exercise is to integrate Skype presence and instant messaging into the help desk application. The starter solution is built with AngularJS and is already configured to authenticate against Azure AD and display the user's profile picture using the Microsoft Graph. Although it helps, you do not need prior experience with AngularJS, basic JavaScript and HTML skills will suffice.

In the first task, you will familiarize yourself with the starter solution and get it running locally.

1. Open a command prompt and browse to the module's **Source > Begin** folder.

2. Open the solution by typing `code .`

		code .

3. The starter solution is considered an AngularJS single page application (SPA) because it uses single index.html page to host all the content. It leverages a Model/View/Controller model to load dynamic content in the single page. The list below lists some of the significant components of the starter solution.

	- **index.html**: the single html that will host all the applications content.
	- **app**: the folder containing all of the application logic and partial views/templates.
		- **templates**: contains all the HTML partial views for the application.
		- **app.js**: defines the primary Angular module for the application and the application routes.
		- **factory.js**: defines an Angular factory that provides properties and services across the application.
		- **controllers.js**: defines all the controllers for the application.
	- **lib**: the folder containing all the frameworks/dependent scripts (ex: Bootstrap, Angular, etc). All of these were imported using **bower**.

4. You should be given **two Office 365 accounts** for this exercise. You should identify one as the **Skype User** and one as the **Web User**. The Skype User will sign into **Skype for Business** and the Web User will use the web application built in this exercise. Open the **factory.js** file in the **app** folder. Go to **lines 32-33** and update the **tenantDomain** and **skypeTestUser** with the tenant domain and Skype User respectively. The example below updates these settings with the contoso tenant and john user.

 ```JavaScript
        31		//Hack...will use hard-coded tickets for demo purposes
        32		var tenantDomain = "contoso.onmicrosoft.com"; //CHANGE THIS TO YOUR OFFICE 365 TENANT
        33		var skypeTestUser = "john@contoso.onmicrosoft.com"; //CHANGE THIS TO THE USER THAT WILL TEST FROM SKYPE
 ```

5. Open the **app.js** file in the **app** folder. Notice how each route includes an attribute to determine if authentication is required or not. This is enabled through **ADAL-Angular**, which is the **Azure Active Directory Authentication Library** (**ADAL**) that manages authentication for the application.

 ```JavaScript
        $routeProvider.when("/login", {
            controller: "loginCtrl",
            templateUrl: "/app/templates/view-login.html",
            requireADLogin: false
        }).when ("/tickets", {
            controller: "ticketsCtrl",
            templateUrl: "/app/templates/view-tickets.html",
            requireADLogin: true
        }).otherwise({
            redirectTo: "/login"
        });
 ```

6. Locate the section of **app.js** where the **ADAL** settings are configured and update the **tenant** property with the tenant domain you are using in Office 365.

 ```JavaScript
        adalProvider.init({
            instance: "https://login.microsoftonline.com/",
            tenant: "TENANT.onmicrosoft.com", //TODO: CHANGE THIS TO YOUR OFFICE 365 TENANT DOMAIN
            clientId: "6fd45769-7a1e-4dc5-a876-90fa781b3d3e",
            endpoints: {
                "https://webdir.online.lync.com": "https://webdir.online.lync.com",
                "https://graph.microsoft.com": "https://graph.microsoft.com"
            }
        }, $httpProvider);
  ```

7. Open the **controllers.js** file in the **app** folder and locate the **loginCtrl**. Notice it's use of **adalSvc** to check if the user is authenticated.

 ```JavaScript
        .controller("loginCtrl", ["$scope", "$location", "adalAuthenticationService", function($scope, $location, adalSvc) {
            if (adalSvc.userInfo.isAuthenticated) {
                $location.path("/tickets");
            }

            $scope.login = function() {
                adalSvc.login();  
            };
        }])
	```

8. Return to the command prompt and type `superstatic --port 8000`. This will start a simple web server to host your client-side web application.

 ```
		superstatic --port 8000
 ```

9. Open a browser and navigate to **http://localhost:8000**. The site should direct you to the **login** view and prompt you to sign-in with **Office 365**.

	![Sign-in with Office 365](Images/Mod4_signin.png?raw=true "Sign-in with Office 365")

	 _Sign-in with Office 365_

10. Sign into the web application using the credentials of the Web User. Once you sign-in, the application should display a list of help desk tickets. In the next Task, you will modify this view to display Skype presence for each user.

	![Help desk tickets](Images/Mod4_tickets.png?raw=true "Help desk tickets")

	 _Help desk tickets_

<a name="Ex2Task2"></a>
#### Task 2 - Sign-in and Presence with Skype for Business ####

In this task, you will introduce the Skype Web SDK into the solution and use it to subscribe and display presence for users in the help desk application. Applications build against the Skype Web SDK need to be registered in Azure AD. We will demonstrate this process, but you will use a predefined application ID in this lab.

1. Create a **skype.js** file in the **app** folder of the solution and populate it with base scaffolding using the **o365-skypefactory** code snippet. This creates a **skype.services** Angular module and populates it with some of the core settings to integrate the **Skype Web SDK**. The **skypeSvc** factory will provide persistent objects and services across all controllers of the application. It uses a **singleton** pattern for defining properties a functions.

 (Code Snippet - _o365-skypefactory_)

 ```JavaScript
        angular.module("skype.services", [])
        .factory("skypeSvc", ["$rootScope", "$http", "$q", function($rootScope, $http, $q) {
            var skypeSvc = {};

            //private properties
            var apiManager = null;
            var client = null;

            //config settings for the app
            skypeSvc.config = {
                apiKey: "a42fcebd-5b43-4b89-a065-74450fb91255", // SDK DF
                apiKeyCC: "9c967f6b-a846-4df2-b43d-5167e47d81e1", // SDK+CC DF
                initParams: {
                    auth: null,
                    client_id: "6fd45769-7a1e-4dc5-a876-90fa781b3d3e", //Client ID of app in Azure AD
                    cors: true,
                    origins: ["https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root"],
                    redirect_uri: "/auth.html",
                    version: "sdk-samples/1.0.0" // this helps to identify telemetry generated by the samples
                }
            };

            //Add additional Skype logic here...signin, status, conversations, etc

            return skypeSvc;
        }]);
```

2. Take note of the redirect_uri property above that is set the /auth.html. This is a page the Skype Web SDK will open in a hidden iFrame to help with token acquisition.

3. Next, open the **index.html** file in the root of the project and add a script reference to the **Skype Web SDK** below the bootstrap reference (you can also use the **o365-skyperef** code snippet for this).

 ```HTML
        <!-- JQuery and Bootstrap references -->
        <script type="text/javascript" src="lib/jquery/dist/jquery.min.js"></script>
        <script type="text/javascript" src="lib/bootstrap/dist/js/bootstrap.min.js"></script>

        <!-- Skype reference -->
        <script src="https://swx.cdn.skype.com/shared/v/latest/SkypeBootstrap.min.js"></script>
 ```

4. You also need to add a reference to the **skype.js** Angular factory you created in step 1. Add that in the app scripts section between the **factory.js** and the **controllers.js**.

 ```HTML
        <!-- App scripts -->
        <script type="text/javascript" src="app/factory.js"></script>
        <script type="text/javascript" src="app/skype.js"></script>
        <script type="text/javascript" src="app/controllers.js"></script>
        <script type="text/javascript" src="app/app.js"></script>
 ```

5. Next, you need to inject a dependency of the **skype.services** module you created in step 1 in the main Angular module of the application. Open the app.js file and modify the helpdesk module definition as follows.

 ```JavaScript
		angular.module("helpdesk", ["skype.services", "helpdesk.services", "helpdesk.controllers", "ngRoute", "AdalAngular"])
 ```

6. Return to the **skype.js** file you created in Step 1. In the "Add additional Skype logic here" section, add an **ensureClient** function on the **skypeSvc** object to initialize the Skype Web SDK. You can follow the code below or use the **o365-skypeEnsureClient** code snippet.

 (Code Snippet - _o365-skypeEnsureClient_)

 ```JavaScript
        //ensures the skype client object is initialized
        var ensureClient = function() {
            var deferred = $q.defer();

            if (client != null)
                deferred.resolve();
            else {
                Skype.initialize({
                    apiKey: skypeSvc.config.apiKeyCC
                }, function (api) {
                    apiManager = api;
                    client = apiManager.UIApplicationInstance;
                    client.signInManager.state.changed(function (state) {
                        $rootScope.$broadcast("stateChanged", state);
                    });
                    deferred.resolve();
                }, function (er) {
                    deferred.resolve(er);
                });
            }

            return deferred.promise;
        };
 ```

7. The code in Step 6 ensures the Skype Web SDK is initialized, but you also need to ensure the user is signed in. Below ensureClient, add an **ensureSignIn** function to the **skypeSvc** object by using the **o365-skypeEnsureSignIn** code snippet. Notice that it checks the uses the **client.signInManager** to check the sign-in state and calls **signIn** if needed.

 (Code Snippet - _o365-skypeEnsureSignIn_)

 ```JavaScript
        //signs into skype
        skypeSvc.ensureSignIn = function() {
            var deferred = $q.defer();

            ensureClient().then(function() {
                //determine if the user is already signed in or not
                if (client.signInManager.state() == "SignedOut") {
                    client.signInManager.signIn(skypeSvc.config.initParams).then(function (z) {
                        //listen for status changes
                        client.personsAndGroupsManager.mePerson.status.changed(function (newStatus) {
                            console.log("logged in status: " + newStatus);
                        });

                        //In the future we can listen for new inbound conversations like this
                        client.conversationsManager.conversations.added(function (conversation) { });

                        //resolve the promise
                        deferred.resolve();
                    }, function (er) {
                        deferred.reject(er);
                    });
                }
                else {
                    //resolve the promise
                    deferred.resolve();
                }
            }, function(er) {
                deferred.reject(er);
            });

            return deferred.promise;
        }
 ```

8. Also notice that the code snippet above includes a line to listen for new inbound conversations. We won't fully implement that in this lab, so this is show just as a future reference.

 ```JavaScript
        //In the future we can listen for new inbound conversations like this
        client.conversationsManager.conversations.added(function (conversation) { });
 ```

9. Next, you should create a **subscribeToStatus** function on the **skypeSvc** to check the status of user that is passed into it (queried using **client.personsAndGroupsManager.createPersonSearchQuery**). The function should also subscribe to status changes for the user. However, you don't want to subscribe the same user more than once, so keep track of subscriptions in a **userSubs** array. You can add this script block using the **o365-skypeSubscribeUser** code snippet.

 (Code Snippet - _o365-skypeSubscribeUser_)

 ```JavaScript
        //subscribes to the status of a user
        var userSubs = [];
        skypeSvc.subscribeToStatus = function(id) {
            var deferred = $q.defer();

            //query for the user by their id
            var query = client.personsAndGroupsManager.createPersonSearchQuery();
            query.text(id);
            query.limit(1);
            query.getMore().then(function (items) {
                //ensure results came back
                if (items.length > 0)
                {
                    //assume the first match is the user
                    var person = items[0].result;
                    person.status.get().then(function (s) {
                        deferred.resolve(s);
                    });

                    //check if we have already subscribed to this user
                    var subMatch = null;
                    for (var i = 0; i < userSubs.length; i++) {
                        if (userSubs[i].id === id) {
                            subMatch = userSubs[i];
                            break;
                        }
                    }
                    if (!subMatch) {
                        //no subscription exists for this user, so create one
                        userSubs.push({ id: id, person: person });

                        //listen for status changes
                        person.status.changed(function(s) {
                            //broadcast the status change to listeners
                            $rootScope.$broadcast("statusChanged", { user: person, status: s });
                        });

                        //subscribe to the status changes
                        person.status.subscribe();
                    }
                }
                else
                    deferred.reject("No matches found");
            });

            return deferred.promise;
        };
 ```

10. Open the **controllers.js** file in the app folder and locate the **ticketsCtrl**. Update it dependency inject the **skypeSvc** you updated in the previous steps.

 ```JavaScript
        .controller("ticketsCtrl", ["$scope", "helpdeskSvc", "skypeSvc", function($scope, helpdeskSvc, skypeSvc) {
            //get the helpdesk tickets
            helpdeskSvc.getTickets().then(function(tickets) {
                $scope.tickets = tickets;
                
                //add o365-skypeSubscribeTickets snippet here
            });

            //get the user's profile picture
            $scope.pic = "/content/nopic.jpg";
            helpdeskSvc.getProfilePic().then(function(img) {
                $scope.pic = img;
            });
            
            //add o365-skypeListenStatus snippet here

            //add o365-skypeCanChat snippet here

            //add o365-skypeStartChat snippet here

            //add $scope.closeChatWindow code here
        }]);
```

11. Next, locate where **getTickets** returns tickets in the "then" promise. After setting $scope.tickets, you should ensure the user is signed into Skype for Business (using **skypeSvc.ensureSignIn**) and then loop through each ticket, subscribing to the ticket opener's status in Skype for Business (using  **skypeSvc.subscribeToStatus**). You can use the o365-skypeSubscribeTickets snippet for this where it is indicated in comments of the starter file.

 ```JavaScript
        //get the helpdesk tickets
        helpdeskSvc.getTickets().then(function(tickets) {
            $scope.tickets = tickets;

            //ensure the user is signed into skype
            skypeSvc.ensureSignIn().then(function() {
                angular.forEach($scope.tickets, function(ticket, index) {
                    //look up the status of the user and listen for changes
                    skypeSvc.subscribeToStatus(ticket.created_by.email).then(function(status) {
                        ticket.created_by.status = status;
                        if (!$scope.$$phase)
                            $scope.$apply(); //this provides async ui update out of thread
                    });
                });
            });
        });
 ```

12. You also want to "listen" for status changes by Skype users. You already subscribed to these events with the Skype Web SDK in step 9, but you need to listen to the **statusChanged** event broadcast from the **skypeSvc**. You can do that as follows or using the **o365-skypeListenStatus** code snippet.

 (Code Snippet - _o365-skypeListenStatus_)

 ```JavaScript
        //listen for status changes
       $scope.$on("statusChanged", function(evt, data) {
           var id = data.user.id().replace("sip:", "").toLowerCase();

           //find all instances of this user
           angular.forEach($scope.tickets, function(ticket, index) {
               if (ticket.created_by.email.toLowerCase() === id) {
                   ticket.created_by.status = data.status;
                   if (!$scope.$$phase)
                       $scope.$apply();
               }
           });
       });
 ```

13. Finally, you need to modify the **view-tickets.html** file located in the **app** > **templates** folder to display the Skype presence of each user. Locate the the repeated table row and update the first table cell as follows or using the **o365-skypePresence** code snippet.

 (Code Snippet - _o365-skypePresence_)

 ```HTML
        <tbody>
            <tr ng-repeat="ticket in tickets">
                <td>
                    <span class="badge" ng-class="ticket.created_by.status"><span class="glyphicon glyphicon-minus"></span></span>
                    <span>{{ticket.created_by.name}}</span>
                </td>
                <td>{{ticket.title}}</td>
                <td>{{ticket.status}}</td>
            </tr>
        </tbody>
 ```

14. Launch the Skype for Business client by opening the Windows start menu and typing **Skype for Business 2016**. Once it launches, sign-in with the user account you have designated as the **Skype User**.

15. Open a browser and navigate to **http://localhost:8000**. After you sign-in with the account you designated as the **Browser User**, you should see a presence icon next to each ticket opener. It might take a few seconds, but the **TEST ACCOUNT** should light up with the correct presence from Skype for Business.

16. Try changing the presence of the Skype User in the Skype for Business 2016 client. After a few seconds, the new presence should display in the helpdesk app.

	![Presence 1](Images/Mod4_presence1.png?raw=true "Presence 1")

	 _Presence 1_

18. You have successfully integrated Skype for Business presence into an existing web application. In the next task you will integrate instant messaging, which the Skype Web SDK makes really easy with a conversation UI.

	![Presence 2](Images/Mod4_presence2.png?raw=true "Presence 2")

	 _Presence 2_

<a name="Ex2Task3"></a>
#### Task 3 - Integrating Instant Messaging ####

In this task, you will continue to customize the Help Desk application to include instant messaging with ticket openers. The Skype for Business Web SDK include power controls that can help deliver a consistent Skype experience in your web applications.

1. Open the **skype.js** file created in the previous task and update it with a new **startConversation** function on the **skypeSvc** object that accepts a SIP address for a user an initiates a conversation with them using the Skype Web SDK. All you have to do to initiate the Skype conversation UI is to use the **apiManager** and the **renderConversation** function, providing a **DIV control** in the page that it can render in ("chatWindowInner" in the example below) and the conversation details (participants, modalities, etc). You can follow the code below or use the **o365-skypeStartConversation** code snippet.

 (Code Snippet - _o365-skypeStartConversation_)

 ```JavaScript
        //start a conversation with a user
        skypeSvc.startConversation = function(sip) {
            //hide all containers
            var containers = document.getElementById("chatWindowInner").children;
            for (var i = 0; i < containers.length; i++) {
                containers[i].style.display = "none";
            }

            var chatSip = sip;
            var uris = [chatSip];
            var container = document.getElementById(chatSip);
            if (!container) {
                //this is a new conversation...create the window
                container = document.createElement("div");
                container.id = chatSip;
                document.getElementById("chatWindowInner").appendChild(container);
                var promise = apiManager.renderConversation(container, { modalities: ["Chat"], participants: uris });
            }
            else
                container.style.display = "block";
        };
 ```

2. Next, you need to update the **view-tickets.html** file in the **app** > **templates** folder to accommodate the conversation UI. Add the following HTML to the bottom of this file or use the **o365-skypeConversationUI** code snippet.

 (Code Snippet - _o365-skypeConversationUI_)

 ```HTML
        <div id="chatWindow" ng-class="{'show': showChatWindow}">
            <span class="glyphicon glyphicon-remove close-chat" ng-click="closeChatWindow()"></span>
            <div id="chatWindowInner"></div>
        </div>
 ```

3. While you are in the **view-tickets.html** file, you should also update the presence indicator to have a click event to start a conversation. In AngularJS, click events are configured using **ng-click** attribute. Set **ng-click** to call **startChat(ticket)**.

 ```HTML
		<span class="badge" ng-class="ticket.created_by.status" ng-click="startChat(ticket)"><span class="glyphicon glyphicon-minus"></span></span>
 ```

4. Next, open the **controllers.js** file in the **app** folder and locate the **ticketsCtrl**. At the bottom of this controller add a private **canChat** function that returns true/false based on the status and it's ability to accept instant messages. You can also add this using the **o365-skypeCanChat** code snippet.

 (Code Snippet - _o365-skypeCanChat_)

 ```JavaScript
        //helper function to check if a status can perform chat
        var canChat = function(status) {
            var chattableStatus = {
                Online: true, Busy: true, Idle: true, IdleOnline: true, Away: true, BeRightBack: true,
                DoNotDisturb: false, Offline: false, Unknown: false, Hidden: false };
            return chattableStatus[status];
        };
 ```

5. Next, define a **startChat** function on the **$scope** object to initiate a conversation with the ticket opener. The function should accept a **ticket parameter** and check if the ticket opener is available to chat based on their availability (via the **canChat** function you just created). If the ticket opener is available for chat, you should call the **skypeSvc.startConversation** function from step 1 of this task. You can use the **o365-skypeStartChat** code snippet or follow the code below.

 (Code Snippet - _o365skypeStartChat_)

 ```JavaScript
        //starts a chat
        $scope.startChat = function(ticket) {
            if (canChat(ticket.created_by.status)) {
                skypeSvc.startConversation("sip:" + ticket.created_by.email);
                $scope.showChatWindow = true;
            }
        };
 ```

6. Finally, add a **closeChatWindow** function on the $scope object to close the chat window. The **$scope.showChatWindow** property will dictate if chat window will be displayed. It is used in-conjunction with the **ng-show** attribute on the chat window markup.

 ```JavaScript
        //closes the chat window
        $scope.closeChatWindow = function() {
            $scope.showChatWindow = false;
        };
 ```

7. Open a browser and navigate to **http://localhost:8000**. After you sign-in with the account you designated as the **Browser User**, you should see a presence icon next to each ticket opener. It might take a few seconds, but the **TEST ACCOUNT** should light up with the correct presence from Skype for Business. Click on the presence icon to initiate a conversation with that user. The Skype conversation UI should fly in from the right and allow you to have a instant messaging conversation with the user.

	![Integrated IM](http://i.imgur.com/i7MxUNC.png)

	 _Integrated IM_

8. By successfully integrated presence and IM into the Help Desk application you have completed this exercise. Know that there are other Skype modalities in preview that you can develop against, including audio and video. Visit the Skype Quick Starts for more information.

<a name="Summary" />
## Summary ##

By completing this module, you should have:


- Skype-enabled and existing web application using the Skype Web SDK.

> **Note:** You can take advantage of the [Visual Studio Dev Essentials]( https://www.visualstudio.com/en-us/products/visual-studio-dev-essentials-vs.aspx) subscription in order to get everything you need to build and deploy your app on any platform.
