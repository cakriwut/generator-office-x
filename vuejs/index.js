'use strict';

const Generator = require('yeoman-generator');
const officeYeoman = require('generator-office/generators/app');

const commandExists = require('command-exists').sync;
const chalk = require('chalk');

const fs = require('fs');

module.exports = class extends Generator {

    constructor(args,opts){
        super(args,opts);

        //this.context = opts.context || {};
    }

    initializing() {
        this.composeWith(
            require.resolve(`generator-office/generators/app`), 
            // {
            //     'skip-install': true,
            //     'projectType': 'jquery',
            //     'js':true,
            //     'ts': false,
            //     'name': 'test-addin',
            //     'host' : 'excel'
            // }      
            {
                'details' : true
            }  
        )
    }

    prompting() {}

    configuring() {}

    writing() {}

    install() {

    }

    end() {

    }
}