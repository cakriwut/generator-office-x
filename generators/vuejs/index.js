'use strict';

const Generator = require('yeoman-generator');
const chalk = require('chalk');
const _ = require('lodash');
const parser = require('camaro');

/* Read from generator-office */
/* eslint-disable camelcase */
const starterCode_1 = require('generator-office/generators/app/config/starterCode');

module.exports = class extends Generator {
  
  writing() {
    // Write, before original Office:App here.
    const done = this.async();
    this._readProjectConfiguration()
      .then(this._copyProjectFiles())
      .then(done())
      .catch(err => {
        this.log(err);
        process.exitCode = 1;
      });
  }

  /* eslint-disable no-negated-condition */
  _readProjectConfiguration() {
    return new Promise((resolve, reject) => {
      try {
        this.log('----------------------------------------------------------------------------------\n');
        this.log(`                      ${chalk.bold.green('Office-X')} custom configuration         \n`);
        this.log('----------------------------------------------------------------------------------\n\n');
        /* Read generator-office files */
        // const packageJson = this.fs.readJSON('package.json', {});
        const manifestXml = this.fs.read('manifest.xml', '<?xml version="1.0" ?>');

        const template = {
          projectId: '/OfficeApp/Id',
          name: '/OfficeApp/DisplayName/@DefaultValue',
          host: '/OfficeApp/Hosts/Host/@Name'
        };

        this.project = parser(manifestXml, template);
        let fileIsTypeScript = this.fs.exists('src/index.ts');

        this.project.scriptType = fileIsTypeScript ? 'Typescript' : 'Javascript';
        this.project.language =  fileIsTypeScript ? 'ts' : 'js';

        this.project.folder = this.project.name;
        /* Set folder if to output param  if specified */

        if (this.options.output !== null) {
          this.project.folder = this.options.output;
        }

        this.project.projectInternalName = _.kebabCase(this.project.name);
        this.project.projectDisplayName = this.project.name;
        this.project.hostInternalName = this.project.host;
        return resolve();
      } catch (err) {
        this.log(err);
        return reject(err);
      }
    });
  }

  _copyProjectFiles() {
    return new Promise((resolve, reject) => {
      try {
        const starterCode = starterCode_1.default(this.project.host);
        const templateFills = Object.assign({}, this.project, starterCode);

        /* Cleanup redundant file */
        if(this.project.language === 'js' && this.fs.exists('src/index.js')) {
          this.fs.delete('src/index.js');
        } else if(this.fs.exists('src/index.ts'))
        {
          this.fs.delete('src/index.ts');
        }

        /* Overwrite office */
        this.fs.copyTpl(this.templatePath(`${this.project.language}/**`), '', templateFills, {
          globOptions: { ignore: `**/*.placeholder` }
        });

        /* Copy all dot files */
        this.fs.copy(this.templatePath(`${this.project.language}/**/.*`),'');

        return resolve();
      } catch (err) {
        this.log(err);
        return reject(err);
      }
    });
  }
};
