import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { enableProdMode } from '@angular/core';
import { AppModule } from './app/app.module';
if (process.env.ENV === 'production') {
  enableProdMode();
}

//bootstrap with Office.js
  Office.initialize = () => {
      console.log("Office init: bootstrapping Angular2");
      platformBrowserDynamic().bootstrapModule(AppModule);
  }