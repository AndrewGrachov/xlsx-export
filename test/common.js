process.env.NODE_ENV = process.env.NODE_ENV || 'test';
process.env.APP_ENV = process.env.APP_ENV || 'development';
global.sinon = require('sinon');
global.chai = require('chai');
global.expect = chai.expect;
global.should = chai.should();
chai.config.includeStack = true;
chai.use(require('sinon-chai'));
require('mocha');