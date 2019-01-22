'use strict';
const Generator = require('yeoman-generator');
const chalk = require('chalk');
const yosay = require('yosay');
const commandExists = require('command-exists');
const fs = require('fs');

module.exports = class extends Generator {

  constructor(args, opts){
    super(args,opts);
    this.argument('projectType', { type: String, required: false });
    this.argument('name', { type: String, required: false });
    this.argument('host', { type: String, required: false });
    this.option('js',{
      type: Boolean,
      require: false,
      desc: 'Project uses JavaScript instead of TypeScript.'
    });
    this.option('ts', {
      type: Boolean,
      required: false,
      desc: 'Project uses TypeScript instead of JavaScript.'
    });
    this.option('output', {
        alias: 'o',
        type: String,
        required: false,
        desc: 'Project folder name if different from project name.'
    });
    this.option('details', {
        alias: 'd',
        type: Boolean,
        required: false,
        desc: 'Get more details on Yo Office arguments.'
    });
  }

  initializing() {
    if (this.options.details) {
      this._detailedHelp();
    }
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
        message: `Choose a ${chalk.bold.green('Extended')} project type or (original) :`,
        choices:[
          { name: 'Office generator (original)', value: 'standard'},
          { name: 'Office Add-in project using Vue framework', value: 'vuejs'}
        ],
        default: 'vuejs'
      }
    ];

    return this.prompt(prompts).then(props => {
      // To access props later use this.props.someAnswer;
      this.props = props;
    });
  }

  configuring() {
    // We need following statement to removed undefined flowing to subgenerator.
    let options = JSON.parse(JSON.stringify(this.options)) || {};
    switch(this.props.extProjectType) {
      case 'vuejs':
        this.composeWith('office-x:vuejs', options, { 
          local: require.resolve('../vuejs')
        });
        break;
      default:
        this.composeWith('office:app', options);  
        break;
    }
  }

  default() {
  }
  writing() {}

  install() {}

  end() {}

  _detailedHelp () {
    this.log(`\nYo ${chalk.underline.bold.greenBright('Office-X')} ${chalk.bgGreen('Arguments')} and ${chalk.bgMagenta('Options.')}\n`);
    this.log(`NOTE: ${chalk.bgGreen('Arguments')} must be specified in the order below, and ${chalk.bgMagenta('Options')} must follow ${chalk.bgGreen('Arguments')}.\n`);
        this.log(`  ${chalk.bgGreen('projectType')}:Specifies the type of project to create. Valid project types include:`);
        this.log(`    ${chalk.yellow('angular:')}  Creates an Office add-in using Angular framework.`);
        this.log(`    ${chalk.yellow('excel-functions:')} Creates an Office add-in for Excel custom functions.  Must specify 'Excel' as host parameter.`);
        this.log(`    ${chalk.yellow('jquery:')} Creates an Office add-in using Jquery framework.`);
        this.log(`    ${chalk.yellow('manifest:')} Creates an only the manifest file for an Office add-in.`);
        this.log(`    ${chalk.yellow('react:')} Creates an Office add-in using React framework.`);
        this.log(`    ${chalk.yellow('vue:')} Creates an Office add-in using Vuejs framework. [${chalk.underline.bold.greenBright('Office-X')}] \n`);
        this.log(`  ${chalk.bgGreen('name')}:Specifies the name for the project that will be created.\n`);
        this.log(`  ${chalk.bgGreen('host')}:Specifies the host app in the add-in manifest.`);
        this.log(`    ${chalk.yellow('excel:')}  Creates an Office add-in for Excel. Valid hosts include:`);
        this.log(`    ${chalk.yellow('onenote:')} Creates an Office add-in for OneNote.`);
        this.log(`    ${chalk.yellow('outlook:')} Creates an Office add-in for Outlook.`);
        this.log(`    ${chalk.yellow('powerpoint:')} Creates an Office add-in for PowerPoint.`);
        this.log(`    ${chalk.yellow('project:')} Creates an Office add-in for Project.`);
        this.log(`    ${chalk.yellow('word:')} Creates an Office add-in for Word.\n`);
        this.log(`  ${chalk.bgMagenta('--output')}:Specifies the location in the file system where the project will be created.`);
        this.log(`    ${chalk.yellow('If the option is not specified, the project will be created in the current folder')}\n`);
        this.log(`  ${chalk.bgMagenta('--js')}:Specifies that the project will use JavaScript instead of TypeScript.`);
        this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
        this.log(`  ${chalk.bgMagenta('--ts')}:Specifies that the project will use TypeScript instead of JavaScript.`);
        this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
        this._exitProcess();
  }

  _exitProcess() {
    process.exit();
  }
};
