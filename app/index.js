'use strict';
const Generator = require('yeoman-generator');
const chalk = require('chalk');
const yosay = require('yosay');
const commandExists = require('command-exists');
const fs = require('fs');

module.exports = class extends Generator {

  constructor(args, opts){
    super(args,opts);
  }

  initializing() {
    this.log(yosay(
      `Welcome to the \n${chalk.bold.green('Extended - Office Add-in')} generator, by ${chalk.bold.green('@cakriwut')}!` +
      `\nBased on \n${chalk.bold.green('Office Add-in generator')}`)
    );
  }

  prompting() {

    const prompts = [
      {
        type: 'list',
        name: 'extProjectType',
        message: `Choose a ${chalk.bold.green('Extended')} project type or (none) :`,
        choices:[
          { name: 'Default (none)', value: 'standard'},
          { name: 'Office Add-in project using Vue framework', value: 'vuejs'}
        ]
      }
    ];

    return this.prompt(prompts).then(props => {
      // To access props later use this.props.someAnswer;
      this.props = props;
    });
  }

  configuring() {
    switch(this.props.extProjectType) {
      case 'vuejs':
        this.composeWith('office-x:vuejs',{}, {
          local: require.resolve('../vuejs')
        });
        break;
      default:
        break;
    }
  }

  writing() {}

  install() {}

  end() {}
};
