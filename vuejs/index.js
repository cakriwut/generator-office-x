'use strict';

const Generator = require('yeoman-generator');
const chalk = require('chalk');
const _ = require('lodash');
const parser = require('camaro');

/* Read from generator-office */
/* eslint-disable camelcase */
const starterCode_1 = require('generator-office/generators/app/config/starterCode');

module.exports = class extends Generator {

    initializing() { }

    prompting() { }

    configuring() { 
        
    }

    default() { }

    writing() { 
        //Write, before original Office:App here.
        const done = this.async();
        this._readProjectConfiguration()
        .then(this._copyProjectFiles())
        .then(done())
        .catch((err)=> {
            this.log(err);
            process.exitCode = 1;
        });       
    }

    install() { }
   
    end() {}

    _readProjectConfiguration() {
        return new Promise((resolve,reject) => {
            try {
                this.log('----------------------------------------------------------------------------------\n');
                this.log(`                      ${chalk.bold.green('Office-X')} custom configuration         \n`);
                this.log('----------------------------------------------------------------------------------\n\n');
                /* Read generator-office files */
                let packageJson = this.fs.readJSON('package.json',{});
                let manifestXml = this.fs.read('manifest.xml','<?xml version="1.0" ?>');

                let template = {
                    projectId : '/OfficeApp/Id',
                    name: '/OfficeApp/DisplayName/@DefaultValue',
                    host: '/OfficeApp/Hosts/Host/@Name'
                }

                this.project = parser(manifestXml,template);

                this.project.scriptType = packageJson.devDependencies.typescript !== null ? 'Typescript' : 'Javascript';
                this.project.language = packageJson.devDependencies.typescript !== null ? 'ts' : 'js';
                
                this.project.folder = this.project.name;
                /* Set folder if to output param  if specified */
                if (this.options.output !== null) {
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
                /* Overwrite office */
                this.fs.copyTpl(this.templatePath(`**`),"", templateFills, { globOptions: { ignore: `**/*.placeholder` } });
                
                return resolve();
            } 
            catch(err){
                this.log(err);
                return reject(err);
                
            }
        });

    }
}