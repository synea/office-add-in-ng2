import 'core-js/es6';
import 'core-js/es7/reflect';
require('zone.js/dist/zone');

//office.js sets replaceState and pushState null which wil break the Angular 2 router. 
window.history.replaceState = function () { };
window.history.pushState = function () { };

if (process.env.ENV === 'production') {
  // Production
} else {
  // Development
  Error['stackTraceLimit'] = Infinity;
  require('zone.js/dist/long-stack-trace-zone');
}
