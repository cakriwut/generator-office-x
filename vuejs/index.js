'use strict';

const Generator = require('yeoman-generator');
const officeYeoman = require('generator-office/generators/app');

const commandExists = require('command-exists').sync;
const chalk = require('chalk');

const fs = require('fs');

module.exports = class extends Generator {

    constructor(args,opts){
        super(args,opts);

    }

    initializing() { }

    prompting() { }

    configuring() { }

    default() { }

    writing() {
        //Write, before original Office:App here.
    }

    install() {
        //Overwrite installation, after original Office:App here.
    }

   
    end() {
        
    }
}