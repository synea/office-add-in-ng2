Starter Poroject for Office Add-in Development with Angular 2
====================

Welcome to my starter project for development of Office Add-ins with Angular 2.
Office Add-ins APIs are based on Javascript, take advantage of several related technologies and sometimes understanding how to use them together 
might be confusing. Here I use ones fit for my needs, yet I try to share links pointing to information regarding others.
Adding Angular 2 with Webpack doesn't make things easier, but I believe provides best environment for integration with future technologies and cloud solutions.

> There are some files needed for this project to work but not included due to security reasons, such as .PFX file to make Express 4 server publish with https.
Some other files need to be changed according to your need, such as .xml manifest file of Office Add-in. Keep reading for how to handle those.

## Quick Start
Assuming you've already read / know information in this README, here are the steps needed in order to get up and running. 
If you are in doubt or not debugging locally on Windows read the rest of this document.

1. Create .PFX certificate file from existing localhost certificate (IISExpress), with a pass phrase/password.
2. Copy .PFX file to **`./server/SecretsAndCerts`** and update `pfx` and `passphrase` properties of **`./server/config.json`** file with file name and pass phrase.
3. Install dependencies both on server project and client-side.
    * Run `npm install` under both **`./server`** and **`./client-side`** folders.
4. Go to **`./client-side`** folder and build Angular 2 app.
    * `cd client-side`
    * `npm run build`
5. You now have a **`dist`** folder under **`./client-side`**, copy it to **`./server`** (overwrite existing one)
6. Copy **`Office-Add-In-Manifest.xml`** to your network shared folder, where you keep Office Add-in manifest files.
7. If you haven't done already, add that network shared folder to Office's Trusted Add-in Catalog list.
8. Execute server.
    * `cd server`
    * `npm start`
9. Open a blank Excel workbook, go to **Insert** >> **Store** >> **Shared Folder** and add `Office Add-in Starter Project`
10. To debug client side code, execute **F12Chooser** tool.

## Selected Technologies

* **Angular 2** w. Webpack as in official docs [WEBPACK: AN INTRODUCTION](https://angular.io/docs/ts/latest/guide/webpack.html)
* **Office Add-in** A simple Excel one. API v1.1 (Office 2016).
* **SSL Certificate** IISExpress generated localhost certificate used (more on that below)
* **Express 4** To test it locally we need a server, Express 4 is selected due to future upgrades and SSL compatibility
* **Add-in Manifest File** Kept very simple, can include more use-cases (like Commands) in the future.
* **Handlebars** Actually we wont use server-side templating, but in case I've selected it simply because I don't like Jade's syntax. Read Handlebars section below.
* **Bootstrap (Twitter)** Lets make our app an eye candy. Native Angular 2 implementation [ng2-bootstrap](https://github.com/valor-software/ng2-bootstrap) used.

## Project Structure
Server side and client side projects are separate, meaning you can run each on their own.
However, client side project runs on http, since Office Add-ins require SSL protected connection and running client-side project on its own won't do much good for local development. 
If you'll debug your add-in using **Office 365** or you have a server on the Internet with https connection, you can work only with **client-side** project.

One other reason why I didn't merge two project is, in case I (you) decide in the future to do more than just serving SSL protected content and do CRUD ops with 3rd party server/service.
Not all scenarios let you work with client-side code only, so you may want to make your server-side code be more involved.

For local development a **gulp** script copies build *dist* folder from client-side folder to server folder and then runs SSL secured web server.

## SSL Certificate

To create an SSL certificate for local development you have few options. Because at the time of development I'm using a Windows machine I've selected easiest one: using the certificate 
created by IISExpress. Other options to create a certificate on Mac or Windows check links below.

Below I've shared steps I took to generate SSL certificate and trust it, if you need screenshots to look at, check links at the bottom of this README.

Open **MMC**, under File menu go to *Add/Remove Snap-in...*, select Certificates from list on the left and click *Add*, it will ask you for which level: **Computer Account** is what you need to select.
Leave **Local computer** selected in next screen, click OK. It's then added to list on the right in previous window, click OK to go back MMC's main screen.

First copy IISExpress generated ***localhost*** certificate (under Personal/Certificates) to ***Trusted Root Certification Authorities***, a simple copy/paste op. 
Then we need to create **.PFX** file with password from IISExpress certificate and we will put under server side project so that Node.js can use it. For that, read section 
**Export the certificate as a .PFX file, include all properties and private key** of article [Make your own SSL Certificate for testing and learning](https://blogs.msdn.microsoft.com/benjaminperkins/2014/05/05/make-your-own-ssl-certificate-for-testing-and-learning/) 
 by Benjamin Perkins. **IMPORTANT:** Pass phrase you enter while creating *.PFX* file will be needed later. 

> Certificate file should be under `./server/SecretsAndCerts` folder and related config settings are defined in `./server/config.json` . This config object also contains 
`passphrase` property I've just mentioned above; it shall be the same pass phrase you entered during .PFX file creation.

 Mac users can use OPENSSL or other methods, unfortunately at the time of writing I couldn't find my notes on that, a Google search shall give you info you need.

## Manifest File & How to Use It
You should know by know that you need to have a local network share (or network shared folder in your local machine) to put your manifest files under and tell Office that given location is 
a **Trusted Add-in Catalog**.

* Open a blank document on Excel/Word
* Go to File >> Options
* Trust Center >> Trust Center Settings...
* Trusted Add-in Catalog
* type your shared folder location and click Add, then OK and exit window, a restart of Office app is required

Add-ins then can be added to your document from Office app's ribbon >> **Insert** >> **Store** (select **Shared Folder** instead of Store when Store window is opened.) Once added, they'll be 
available on ribbon >> **Insert** >> **My Add-ins** for quickly executing them.

**NOTE:** Manifest file has an `<Id> </Id> ` tag that contains GUID. It should be unique, so change it from the one provided. You can use Visual Studio's *Create GUID* tool under Tools menu item.
You also need to change sections like `ProviderName`, `DisplayName`, `Description`...etc.

## Debugging
Since this project is based on Node.js tools and not a Visual Studio project, client-side debugging isn't integrated with development IDE. If you are uploading your add-in to Office 365
 to test it, you can use browser dev tools to debug.

For local machine client-side debugging, you are not in luck if you are on a Macintosh computer, try to get an Office 365 account and debug there.
Open up a new workbook in Excel Online. Click on the **INSERT** tab and then **Office Add-ins**. On the Office Add-ins dialog click **Upload my Add-In** and select 
**Office-Add-In-Manifest.xml**. The add-in will now be available on the **HOME** tab. **F12** to debug the add-in in browser. 

On **Windows 10** now there is a great tool: **F12 Chooser**: it's a tool for WinForms applications hosting Web Browser Control. Execute your add-in and just open folder from explorer *C:\Windows\SysWOW64\F12* 
(or *C:\Windows\System32\F12*) and execute F12Chooser.exe
You'll see your add-in name, click that button and a IE dev-tools window will open.

## Handlebars
We are using mostly static html/js files generated by client-side sub-project, but in case we'll need a templating solution Handlebars is here to rescue.
Angular uses curly brackets just like Handlebars for string interpolation so we need to tell Handlebars which double brackets {{ }} is for Angular, which one is for 
himself to process.
Simply use 3 curly brackets `{{{ someVar }}}` for Handlebars (to make distinction) and use slash in front of brackets to cancel Handlebars interpolation and print it in html for Angular to process `\{{ vm.someProp }}`

## Webpack
Depending on other 3rd party frameworks you'll use, some changes may be needed in webpack configuration. That's out of scope of this project.

However, I'd like to point out something, Webpack settings as presented in [WEBPACK: AN INTRODUCTION](https://angular.io/docs/ts/latest/guide/webpack.html) are debugging ready.
Meaning it creates .map files next to uglified/bundled javascript files. That may be something you don't want since they are large and they make your source code open to investigation to others.
To cancel this behavior, open `client-side/config/webpack.prod.js` 

and comment out this line:

`devtool: 'source-map',`

## Resources

Here are some important links

* [Office Add-ins](https://dev.office.com/getting-started/addins)
* [How to trust the IIS Express Self-Signed Certificate](https://blogs.msdn.microsoft.com/robert_mcmurray/2013/11/15/how-to-trust-the-iis-express-self-signed-certificate/)
* [Make your own SSL Certificate for testing and learning](https://blogs.msdn.microsoft.com/benjaminperkins/2014/05/05/make-your-own-ssl-certificate-for-testing-and-learning/)
* [Creating a Secure (TLS) Node.js MVC Application](https://shellmonger.com/2015/07/08/creating-a-secure-node-mvc-application/)
* [F12Chooser: Amazing New Tool for WinForms Application Hosting Web Browser Control](https://blogs.technet.microsoft.com/prakashpatel/2015/11/04/vs2015-remote-debugging-javascript-part-3-f12chooser/)
* [Building an Office add-in with Angular 2 sample in C#, Bash/shell for Visual Studio 2015](https://code.msdn.microsoft.com/office/building-an-office-add-in-a9d506cd)