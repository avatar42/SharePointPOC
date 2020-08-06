# SharepointPOC
Example of top and side nave in single app based on [this set of blog posts](http://codingfix.com/cordova-application-navigation-system/).

**Note this a POC from 2018 and pom.xml needs updating to newer versions of tools if planning to use. See Github alerts.**


Most of the libs and APIs required are included in the repo to make things easier.
Tools I have installed:
Java 1.8 JSK
Android Studio 3.1.2 with path pointed to the assocated SDK with build tools 28-rx2


To get running (might not need all of these)
cd to the SharepointPOC folder
To rebuild the platforms and node_modules folders run 

**npm install material-design-icons**

**cordova plugin add cordova-plugin-whitelist**

Add what ever platform@versions you are testing with.

**cordova platform add android@7.0.0** and or other plaforms ios, windows, browser

**cordova plugin add cordova-plugin-ms-adal**

**cordova plugin add https://github.com/EddyVerbruggen/Toast-PhoneGap-Plugin.git**


Then to build and run

**cordova clean** always a good idea to force a rebuild


**cordova run** to run in browser or **simulate ios** to run on an android device for example. **Note however the adal calls will not work in either case and simulate android only works with pre android 6.3.0. And 6.2.3 seems to have bugs that keep some callbacks passed to adal methods from being used.**

If running the app on an android device or emulator you might also want to run **adb logcat** after the app starts to see the log as messages are written to it.

If attempting to debug android emulator is often fails to attach and you need to 

In Visual Studio Code I've notice Cordova hangs sometimes waiting for a response from **Run Android on emulator**. Especially if the app is not on the phone already or the build tools seem to think it is not. To get around this run **Attach to running android on emulator** after closing the error dialog. Though I've also noticed hangs waiting for emulator ready that are cleared by uninstalling the app from the emulator.

**adb uninstall com.dea42.sharepointpoc**

**adb install platforms\android\app\build\outputs\apk\debug\app-debug.apk**

or

**adb install platforms\android\build\outputs\apk\android-debug.apk**

Depending on your version of Android SDK build tools


Note tons of extra debug "dprint" is enabled due frequent issues getting debugger to work. Comment out the body of function dprint(msg, isError), in www/index.html to disable. 

Also probably going to shelve and go with ionic.
