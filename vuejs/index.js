'use strict';

const Generator = require('yeoman-generator');
const officeYeoman = require('generator-office/generators/app');

const _ = require('lodash');
const commandExists = require('command-exists').sync;
const chalk = require('chalk');
const path = require('path');
const parser = require('camaro');

const fs = require('fs');
const starterCode_1 = require('generator-office/generators/app/config/starterCode');

module.exports = class extends Generator {

    constructor(args,opts){
        super(args,opts);

    }

    initializing() { }

    prompting() { }

    configuring() { 
        
    }

    default() { }

    writing() { 
        //Write, before original Office:App here.
       
    }

    install() {
        //Overwrite installation, after original Office:App here.
        const done = this.async();
        this._readProjectConfiguration()
        .then(this._copyProjectFiles())
        .then(
            done()
        )
        .catch((err)=> {
            this.log(err);
            process.exitCode = 1;
        });
    }

    _writeLog(callback) {
        return new Promise((resolve,reject) => {
            callback();
            return resolve();
        })
    }

   
    end() {}

    _readProjectConfiguration() {
        return new Promise((resolve,reject) => {
            try {
                this.log('----------------------------------------------------------------------------------\n');
                this.log(`                      ${chalk.bold.green('Office-X')} custom configuration         \n`);
                this.log('----------------------------------------------------------------------------------\n\n');
                /* Read generator-office result */
                let workingPath = process.cwd();
                let manifestPath = path.join(workingPath,'manifest.xml');
                let packageJsonPath = path.join(workingPath,'package.json');
                var packageJson = JSON.parse(fs.readFileSync(packageJsonPath,'utf-8'));                
                var data = fs.readFileSync(manifestPath,'utf-8');
                let template = {
                    projectId : '/OfficeApp/Id',
                    name: '/OfficeApp/DisplayName/@DefaultValue',
                    host: '/OfficeApp/Hosts/Host/@Name'
                }

                this.project = parser(data,template);
                
                if(packageJson.devDependencies.typescript != null){
                    this.project.scriptType = 'Typescript';
                    this.project.language = 'ts';
                } else {
                    this.project.scriptType = 'Javascript';
                    this.project.language = 'js';
                }
                
                this.project.folder = this.project.name;
                /* Set folder if to output param  if specified */
                if (this.options.output != null) {
                    this.project.folder = this.options.output;
                }

                this.project.projectInternalName = _.kebabCase(this.project.name);
                this.project.projectDisplayName = this.project.name;
                this.project.hostInternalName = this.project.host;                
                return resolve();
            }
            catch(err) {
                this.log(err);
                return reject(err);
            }
        });
    }
    _copyProjectFiles() {
        return new Promise((resolve,reject) => {
            try {                  
                const starterCode = starterCode_1.default(this.project.host);
                const templateFills = Object.assign({},this.project,starterCode);
                this.destinationRoot(this.destinationPath());
                this.fs.copyTpl(this.templatePath(`**`),this.destinationPath(), templateFills, { globOptions: { ignore: `**/*.placeholder` } });
                
                return resolve();
            } 
            catch(err){
                this.log(err);
                return reject(err);
                
            }
        });

    }
}