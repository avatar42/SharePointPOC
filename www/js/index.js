var authenticationContext;
var testUserId = null;

var useV2 = true; // use newer graph API instead of limited Azure API

var authority = "",
    redirectUri = "",
    resourceUri = "",
    clientId = "",
    graphApiVersion = "",
    accessToken = "";

function switchV2() {
    //Currently having issues
    dprint("Calling switchV2");
    authority = "https://login.microsoftonline.com/common"; //do not add /oauth2/v2.0/authorize";
    redirectUri = "https://citizant.sharepoint.com";
    resourceUri = "https://graph.microsoft.com";
    clientId = "10a2d9f2-3571-4209-be4e-cf65ff348b36";
    linkUrl = "https://graph.microsoft.com/v1.0/sites/root/lists/da7cd400-0cd9-4dde-9167-3049747f195a/items?expand=fields"
    graphApiVersion = "v1.0";
    useV2 = true;
    app.createContext();
    document.getElementById('apiVersion').innerHTML = "Using V2 API";
    dprint("API swtiched to version 2");
}

function switchV1() {
    dprint("Calling switchV1");
    authority = "https://login.windows.net/common";
    redirectUri = "https://citizant.sharepoint.com";
    resourceUri = "https://graph.windows.net";
    clientId = "10a2d9f2-3571-4209-be4e-cf65ff348b36";
    graphApiVersion = "2013-11-08";
    useV2 = false;
    app.createContext();
    document.getElementById('apiVersion').innerHTML = "Using V1 API";
    dprint("API swtiched to version 1");
}

function switchV() {
    app.logout();
    if (useV2) {
        switchV1();
    } else {
        switchV2()
    }
}


// var tenantName = 'citizant.sharepoint.com';
// var endpointUrl = resourceUri + tenantName;

function pre(json) {
    var out = '<pre>' + JSON.stringify(json, null, 4) + '</pre>';
    document.getElementById("rawResponse").innerHTML = out;

    return out;
}

var app = {
    initialize: function () {
        dprint('initialize: ');
        //document.addEventListener('deviceready', this.onDeviceReady.bind(this), false);
        document.addEventListener('deviceready', app.onDeviceReady, false);
        dprint('addEventListener(deviceready)');
        document.getElementById('create-context').addEventListener('click', app.createContext);
        dprint('addEventListener(create-context)');
        document.getElementById('acquire-token').addEventListener('click', app.acquireToken);
        dprint('addEventListener(acquire-token)');
        document.getElementById('acquire-token-silent').addEventListener('click', app.login);
        dprint('addEventListener(acquire-token-silent)');
        document.getElementById('read-tokencache').addEventListener('click', app.readTokenCache);
        dprint('addEventListener(read-tokencache)');
        document.getElementById('clear-tokencache').addEventListener('click', app.logout);
        dprint('addEventListener(clear-tokencache)');

        // function toggleMenu() {
        //     // menu must be always shown on desktop/tablet
        //     if (document.body.clientWidth > 480) return;
        //     var cl = document.body.classList;
        //     if (cl.contains('left-nav')) { cl.remove('left-nav'); }
        //     else { cl.add('left-nav'); }
        // }

        // document.getElementById('sidemenu-open').addEventListener('click', toggleMenu);
        // dprint('addEventListener(sidemenu-open)');

    },
    onDeviceReady: function () {
        app.logArea = document.getElementById("debugLog");
        app.log("Cordova initialized, 'deviceready' event was fired");
        authenticationContext = Microsoft.ADAL.AuthenticationContext;
        // these call createContext()
        if (useV2) {
            switchV2();
        } else {
            switchV1();
        }

        dprint('app.onDeviceReady');
        document.addEventListener("backbutton", onBackKeyDown, false);
        dprint("backbutton event listener installed");
        $('.carousel').carousel('pause');
        document.getElementById('appOpened').innerHTML = '<span class="timestamp">Started: ' + new Date().toLocaleTimeString() + ': </span>';
    },
    // Update DOM on a Received Event
    receivedEvent: function (id) {
        var parentElement = document.getElementById(id);
        // var listeningElement = parentElement.querySelector('.listening');
        // var receivedElement = parentElement.querySelector('.received');

        // listeningElement.setAttribute('style', 'display:none;');
        // receivedElement.setAttribute('style', 'display:block;');

        dprint('Received Event: ' + id);
    },
    log: function (message, isError) {
        isError ? console.error(message) : console.log(message);
        var logItem = document.createElement('li');
        logItem.classList.add("topcoat-list__item");
        isError && logItem.classList.add("error-item");
        var timestamp = '<span class="timestamp">' + new Date().toLocaleTimeString() + ': </span>';
        logItem.innerHTML = (timestamp + message);
        app.logArea.insertBefore(logItem, app.logArea.firstChild);
    },
    error: function (err) {
        app.log(err, true);
        window.plugins.toast.showWithOptions({
            message: err,
            duration: "short", // which is 2000 ms. "long" is 4000. Or specify the nr of ms yourself.
            position: "bottom",
            addPixelsY: -40  // added a negative value to move it up a bit (default 0)
        });
    },
    createContext: function () {
        dprint('Entering createContext')
        // Note is async / deferred 
        authenticationContext.createAsync(authority)
            .then(function (context) {
                app.authContext = context;
                app.log("Created authentication context for authority URL: " + context.authority);
            }, app.logTokenErr);
    },
    acquireToken: function (authCompletedCallback, authFailedCallback) {
        if (app.authContext == null) {
            app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }
        if (!authCompletedCallback) {
            authCompletedCallback = app.logTokenGood;
        }
        if (!authFailedCallback) {
            authFailedCallback = app.logTokenErr;
        }
        dprint('Calling:app.authContext.acquireTokenAsync(' + resourceUri + ',' + clientId + ',' + redirectUri + ') on ' + app.authContext.authority)
        app.authContext.acquireTokenAsync(resourceUri, clientId, redirectUri)
            .then(authCompletedCallback, authFailedCallback);

    },
    storeUserId: function () {
        testUserId = null;
        app.authContext.tokenCache.readItems().then(function (cacheItems) {
            if (cacheItems.length > 0) {
                testUserId = cacheItems[0].userInfo.userId;
                dprint("Setting userId:" + testUserId);
            }
        });
    },
    logSuccess: function (authResult) {
        app.log('Query successful: ' + pre(authResult));
        enterSecondaryView('speaker_notes');
    },
    logTokenGood: function (authResult) {
        app.log('Acquired token successfully: ' + pre(authResult));
        app.storeUserId();
        enterSecondaryView('speaker_notes');
    },
    logTokenErr: function (err) {
        app.error("Failed to acquire token: " + pre(err));
    },
    login: function (authCompletedCallback, authFailedCallback) {
        if (app.authContext == null) {
            app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }
        if (!authCompletedCallback) {
            authCompletedCallback = app.logTokenGood;
        }
        if (!authFailedCallback) {
            authFailedCallback = app.logTokenErr;
        }

        // testUserId parameter is needed if you have > 1 token cache items to avoid "multiple_matching_tokens_detected" error
        // Note: This is for the test purposes only
        testUserId = null;
        app.authContext.tokenCache.readItems().then(function (cacheItems) {
            if (cacheItems.length > 0) {
                testUserId = cacheItems[0].userInfo.uniqueId;
            }
            if (testUserId != null) {
                dprint("Attempting acquireTokenSilentAsync(" + resourceUri + ", " + clientId + ")");
                app.authContext.acquireTokenSilentAsync(resourceUri, clientId).then(authCompletedCallback, function (err) {
                    app.error("Failed to acquire token silently: " + pre(err));
                    dprint("We require user credentials so triggering authentication dialog");
                    app.acquireToken(authCompletedCallback, authFailedCallback);
                });
            } else {
                app.acquireToken(authCompletedCallback, authFailedCallback);
            }
        }, function (err) {
            app.error("Unable to get User ID from token cache. Have you acquired a token already? " + pre(err));
        });
    },
    readTokenCache: function () {
        if (app.authContext == null) {
            app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        app.authContext.tokenCache.readItems()
            .then(function (res) {
                var text = "Read token cache successfully. There is " + res.length + " items stored.";
                if (res.length > 0) {
                    text += "The first one is: " + pre(res[0]);
                }
                app.log(text);

            }, function (err) {
                app.error("Failed to read token cache: " + pre(err));
            });
    },
    logout: function () {
        if (app.authContext == null) {
            app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        app.authContext.tokenCache.clear().then(function () {
            app.log("Cache cleaned up successfully.");
            testUserId = null;
        }, function (err) {
            app.error("Failed to clear token cache: " + pre(err));
        });

        enterSecondaryView('history');
    },
    search: function () {
        dprint("searchUsers");
        document.getElementById('userlist').innerHTML = "";

        app.login(function (authresult) {
            dprint('Entering search.login');
            var url = "";
            var searchText = document.getElementById('searchfield').value;

            if (useV2) {
                // graph.windows.com
                url = resourceUri + "/" + graphApiVersion + "/me/";
            } else {
                // graph.windows.net has no includes or other "like" op. See https://msdn.microsoft.com/en-us/library/azure/ad/graph/howto/azure-ad-graph-api-supported-queries-filters-and-paging-options
                // url = resourceUri + "/" + authResult.tenantId + "/tenantDetails?api-version=" + graphApiVersion;

                url = resourceUri + "/" + authResult.tenantId + "/users?api-version=" + graphApiVersion;

                url = searchText ? url + "&$filter=startswith(displayName,'" + searchText + "')" : url + "&$orderby=displayName&$top=10";
                // url = searchText ? url + "&$filter=mailNickname eq '" + searchText + "'" : url + "&$top=10";
            }
            dprint('Calling requestData(authresult, ' + url + ')');
            app.requestData(authresult, url);
        }, app.logTokenErr);
    },
    // Makes Api call to receive user list.
    requestData: function (authResult, url) {
        dprint("C-Calling:" + url);
        var req = new XMLHttpRequest();

        req.open("GET", url, true);
        req.setRequestHeader('Authorization', 'Bearer ' + authResult.accessToken);
        document.getElementById('rawResponse').innerHTML = "";

        req.onload = function (e) {
            if (e.target.status >= 200 && e.target.status < 300) {
                dprint("C-Response:" + e.target.response);
                document.getElementById('rawResponse').innerHTML = e.target.response;
                var data = JSON.parse(e.target.response);
                dprint("C-data:" + Object.prototype.toString.call(data).slice(8, -1) + ":" + data);
                if (!data) {
                    error("Unable to parse:" + json);
                    return;
                }
                var dataType = data && Object.keys(data)[0];
                dprint("C-dataType:" + Object.prototype.toString.call(dataType).slice(8, -1) + ":" + dataType);

                app.renderUserListData(data);
                return;
            }
            document.getElementById('rawResponse').innerHTML = e.target.response;
            app.error('Data request failed: ' + e.target.response);
        };
        req.onerror = function (e) {
            app.error('Data request failed: ' + e.error);
        }

        req.send();
    },
    // Renders user list.
    renderUserListData: function (data) {
        var users = data && data.value;
        dprint("C-users:" + Object.prototype.toString.call(users).slice(8, -1) + ":" + users);
        if (!users || users.length === 0) {
            app.error("No users found");
            return;
        }

        var userlist = document.getElementById('userlist');
        userlist.innerHTML = "";

        // Helper function for generating HTML
        function $new(eltName, classlist, innerText, children, attributes) {
            var elt = document.createElement(eltName);
            classlist.forEach(function (className) {
                elt.classList.add(className);
            });

            if (innerText) {
                elt.innerText = innerText;
            }

            if (children && children.constructor === Array) {
                children.forEach(function (child) {
                    elt.appendChild(child);
                });
            } else if (children instanceof HTMLElement) {
                elt.appendChild(children);
            }

            if (attributes && attributes.constructor === Object) {
                for (var attrName in attributes) {
                    elt.setAttribute(attrName, attributes[attrName]);
                }
            }

            return elt;
        }

        users.map(function (userInfo) {
            return $new('li', ['topcoat-list__item'], null, [
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'First name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.givenName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Last name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.surname || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'UPN: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.userPrincipalName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Phone: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.telephoneNumber || ''
                    })
                ])
            ]);
        }).forEach(function (userListItem) {
            userlist.appendChild(userListItem);
        });
    }


};

app.initialize();


$(".carousel-inner").swipe({
    swipeLeft: function () {
        dprint('Received Event: swipeLeft');
        $(this).parent().carousel('next');
    },
    swipeRight: function () {
        dprint('Received Event: swipeRight');
        $(this).parent().carousel('prev');
    },
    threshold: 200,
    excludedElements: "label, button, input, select, textarea, .noSwipe"
});

$('#myCarousel').on('slide.bs.carousel', function (e) {
    var page = $('.item.active').attr('id');
    dprint('Received Event: ' + e.direction + ":" + page);

    if (e.direction == 'left') {
        if (page == 'today') {
            dprint('Received Event: swipeLeft block ' + page);
            e.preventDefault();
        }
    } else {
        if (page == 'home') {
            dprint('Received Event: swipeRight block ' + page);
            e.preventDefault();
        }
    }
});

function jump(pageID) {
    dprint('jump:' + pageID);
    $('#myCarousel .carousel-inner div.active').removeClass('active');
    $('#' + pageID).addClass('active');
}


function enterMenu() {
    if (testUserId) {
        dprint("Switching menu item to logout");
        document.getElementById('loginItem').innerHTML = '<a href="javascript:app.logout();"><i data-index="9" class="material-icons">unlock</i> Logout</a>';
        document.getElementById('loginStatus').innerHTML = 'Logged in as:' + testUserId;
    } else {
        dprint("Switching menu item to login");
        document.getElementById('loginItem').innerHTML = '<a href="javascript:app.login();"><i data-index="9" class="material-icons">lock</i> Login</a>';
        document.getElementById('loginStatus').innerHTML = 'Logged out';
    }
    $('#sidemenu').css({ 'display': 'block' }).addClass('animated slideInRight');
    $('#overlay').fadeIn();
    setTimeout(function () {
        $('#sidemenu').removeClass('animated slideInRight');
    }, 1000);
}

function exitMenu() {
    dprint("Entering exitMenu()");

    $('#sidemenu').addClass('animated slideOutRight');
    $('#overlay').fadeOut();
    setTimeout(function () {
        $('#sidemenu').removeClass('animated slideOutRight').css({ 'display': 'none' });
    }, 1000);
    dprint("Leaving exitMenu()");

}

//Secondary views
var secondaryViews = [];
var activeView;
function enterSecondaryView(id) {
    exitMenu();
    if (id == '') {
        return;
    }
    dprint("switching to SecondaryView:" + id);
    activeView = id;
    var animation = 'slideInLeft';
    processSecondaryViewsArrayOnEnter();
    $(document).scrollTop(0);
    doEnterAction();
    $('#' + activeView).css({ 'display': 'block' }).addClass('animated ' + animation);
    setTimeout(function () {
        dprint('setTimeout:' + activeView);
        // it make have done by the time this runs.
        if (activeView == '') {
            return;
        }
        $('#' + activeView).removeClass('animated ' + animation);
        if (secondaryViews.length > 1) {
            $('#' + secondaryViews[secondaryViews.length - 2]).hide();
        }
        activeView = '';
    }, 1000);
}
function processSecondaryViewsArrayOnEnter() {
    if ($.inArray(activeView, secondaryViews) == -1) {
        secondaryViews.push(activeView);
    }
}
function doEnterAction() {
    dprint('doEnterAction:' + activeView);
    if (activeView == 'dashboard') {
        //here your specific stuff
    } else if (activeView == 'speaker_notes') {
        //here your specific stuff
    } else if (activeView == 'search') {
        //here your specific stuff
    } else if (activeView == 'history') {
        //here your specific stuff
    }
}

function exitSecondaryView(id, sender) {
    activeView = id;
    var animation = 'slideOutLeft';
    processSecondaryViewsArrayOnLeave();
    doLeaveAction();
    $('#' + activeView).addClass('animated ' + animation);
    setTimeout(function () {
        $(document).scrollTop(0);
        $('#' + activeView).css('display', 'none').removeClass('animated ' + animation);
        activeView = '';
    }, 1000);
}
function processSecondaryViewsArrayOnLeave() {
    if (secondaryViews.length > 1) {
        $('#' + secondaryViews[secondaryViews.length - 2]).show();
    }
    secondaryViews.pop();
}
function doLeaveAction() {
    dprint('doLeaveAction:' + activeView);
    if (activeView == 'dashboard') {
        //here your specific stuff
    } else if (activeView == 'speaker_notes') {
        //here your specific stuff
    } else if (activeView == 'search') {
        //here your specific stuff
    } else if (activeView == 'history') {
        //here your specific stuff
    }
}


var lastTimeBackPress = 0;
var timePeriodToExit = 2000;
function onBackKeyDown(e) {
    e.preventDefault();
    e.stopPropagation();
    dprint('Received Event: onBackKeyDown');
    if (secondaryViews.length == 0) {
        dprint('sv length is 0');
        var page = $('.item.active').attr('id');
        if (page == 'home') {
            if (new Date().getTime() - lastTimeBackPress < timePeriodToExit) {
                navigator.app.exitApp();
            } else {
                window.plugins.toast.showWithOptions({
                    message: "Press again to exit.",
                    duration: "short", // which is 2000 ms. "long" is 4000. Or specify the nr of ms yourself.
                    position: "bottom",
                    addPixelsY: -40  // added a negative value to move it up a bit (default 0)
                });
                lastTimeBackPress = new Date().getTime();
            }
        } else {
            $('.carousel').carousel('prev');
        }
    } else {
        var activeView = secondaryViews[secondaryViews.length - 1];
        exitSecondaryView(activeView);
    }
}


// element to function links
$(document).on('click', '.dashboard', function (e) {
    e.preventDefault();
    dprint('Received Event: click.dashboard');

    enterSecondaryView('dashboard');
});
$(document).on('click', '.speaker_notes', function (e) {
    e.preventDefault();
    dprint('Received Event: click.speaker_notes');

    enterSecondaryView('speaker_notes');
});
$(document).on('click', '.search', function (e) {
    e.preventDefault();
    dprint('Received Event: click.search');

    enterSecondaryView('search');
});
$(document).on('click', '.history', function (e) {
    e.preventDefault();
    dprint('Received Event: click.history');

    enterSecondaryView('history');
});

$(".secondary-view").swipe({
    swipeLeft: function () {
        dprint('Received Event: .secondary-view.swipeLeft');
        var activeView = $(this).attr('id');
        exitSecondaryView(activeView);
    },
    threshold: 200,
    excludedElements: "label, button, input, select, textarea, .noSwipe"
});

$(".back-button").swipe({
    swipeLeft: function () {
        dprint('Received Event: .back-button.swipeLeft');
        var activeView = $(this).attr('id');
        exitSecondaryView(activeView);
    },
    threshold: 200,
    excludedElements: "label, button, input, select, textarea, .noSwipe"
});

$(document).on('click', '.cfx-topbar .material-icons', function (e) {
    var target = $(this).data('index');
    dprint('Received Event: menu click:' + target + ":" + $(this).text() + ":" + JSON.stringify($(this)) + ":" + $(this).data);
    $('.carousel').carousel(target);
    switch (target) {
        case 1:
            $(this).parent().carousel('home');
            break;
        case 2:
            $(this).parent().carousel('note');
            break;
        case 3:
            $(this).parent().carousel('list');
            break;
        case 4:
            $(this).parent().carousel('today');
            break;
    }
});

// side menu
$(document).on('click', '#sidemenu-open', function (e) {
    e.preventDefault();
    dprint('Received Event: click#sidemenu-open');
    enterMenu();
});
$(document).on('click', '#sidemenu-close', function (e) {
    e.preventDefault();
    dprint('Received Event: click#sidemenu-close');
    exitMenu();
});

// Implements search operations.

$(document).on('click', '#searchButton', function (e) {
    e.preventDefault();
    dprint('Received Event: click#searchButton');
    app.search();
});
dprint("searchButton event listener installed");
