'use strict';
const appInsights = require('applicationinsights');
const Generator = require('yeoman-generator');
const chalk = require('chalk');
const yosay = require('yosay');

const typescript = `Typescript`;
const javascript = `Javascript`;
const originalProjectTypes = ['angular', 'excel-function', 'jquery', 'manifest', 'react'];
const extProjectTypes = ['vue'];
appInsights.setup('fee06a0c-4806-42fc-9ed8-96a2ccf3144d').start();
const insight = appInsights.defaultClient;
delete insight.context.tags['ai.cloud.roleInstance'];
delete insight.context.tags['ai.device.osVersion'];
delete insight.context.tags['ai.device.osArchitecture'];
delete insight.context.tags['ai.device.osPlatform'];
insight.context.tags['ai.cloud.role'] = 'office-x:main';

module.exports = class extends Generator {
  constructor(args, opts) {
    super(args, opts);
    this.argument('projectType', { type: String, required: false });
    this.argument('name', { type: String, required: false });
    this.argument('host', { type: String, required: false });
    this.option('js', {
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

    this.log(
      yosay(
        `Welcome to the \n${chalk.bold.green('Extended - Office Add-in')}generator, by ${chalk.bold.green('@cakriwut')}!\n` +
        `Based on \n${chalk.bold.green('Office Add-in generator')}`
      )
    );
  }

  prompting() {
    // If projectType not specified, or project type not valid list
    const checkOptions = this.options.projectType === null ||
                        !originalProjectTypes.concat(extProjectTypes).includes(this.options.projectType);
    const prompts = [
      {
        type: 'list',
        name: 'extProjectType',
        message: `Choose a ${chalk.bold.green('Extended')} project type or (original):`,
        choices: [
          { name: 'Office generator (original)', value: 'standard' },
          { name: 'Office Add-in project using Vue framework', value: 'vue' }
        ],
        default: 'vue',
        when: checkOptions
      },       
      {
        name: 'scriptType',
        type: 'list',
        message: 'Choose a script type:',
        choices: [javascript,typescript],
        default: javascript,
        when: answers => answers.extProjectType !== 'standard' && this.options.js === undefined && this.options.ts === undefined
      }
    ];

    return this.prompt(prompts)
      .then(answers => {
        this.props = answers;
        const composedOptions = {};
        composedOptions['skip-install'] = true;     
       
        if(this.props.scriptType !== undefined) {
          let resultForScriptType = (this.props.scriptType !== undefined) ? this.props.scriptType : ( this.options.js === undefined ? typescript : javascript);
          this.options.js = (this.options.js === undefined ) ? resultForScriptType === javascript : this.options.js;
          this.options.ts = (this.options.ts === undefined ) ? resultForScriptType === typescript : this.options.ts;
          this.options.scriptType = resultForScriptType;
        }

        // Props.extProjectType from prompt. If extProjectType, then just default to jQuery.
        // Otherwise, just feed to subgenerator office:app
        if (this.props.extProjectType !== null && extProjectTypes.includes(this.props.extProjectType)) 
        {
          this.options.extProjectType = this.props.extProjectType;          
          this.options.projectType = 'jquery';
            /* Temporary */
          this.options.js = true;
          this.options.ts = false;          
        }

        /* Create insights */;
        insight.trackEvent('OfficeX', this.options);

        const options = JSON.parse(JSON.stringify(Object.assign({}, this.options, composedOptions))) || {};
        this.composeWith('office:app', options);
        //this.composeWith(require.resolve('generator-office/generators/app'),options);
      })
      .catch(err => 
      {
        this.log(err);
        //insight.trackException(new Error('Prompting Error: ' + err));
      });
  }

  install() {
    try {
      if (this.options['skip-install']) {
        // this.installDependencies({
        //   npm: false,
        //   bower: false,
        //   callback: this._postInstallMessage.bind(this)
        // });
        this._postInstallMessage();
      } else {        
        this.installDependencies({
          npm: true,
          bower: false,
          callback: this._postInstallMessage.bind(this)
        });        
      }
    } catch (err) {
       insight.trackException(new Error('Install Error: ' + err));
    }
  }

  customize() {
    const options = JSON.parse(JSON.stringify(this.options)) || {};
    insight.trackEvent('Extended_ProjectType', {
      ExtendedProjectType: this.options.extProjectType
    });
    switch (this.options.extProjectType) {
      case 'vue':
        this.composeWith('office-x:vuejs', options, {
          local: require.resolve('../generators/vuejs')
        });
        break;

      default:
        // Default office generator
        break;
    }
  }

  end() {}

  _postInstallMessage() {
            /* Next steps and npm commands */
            this.log('----------------------------------------------------------------------------------------------------------\n');
            this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
            this.log(`      1. Launch your local web server via ${chalk.inverse(' npm start ')} (you may also need to`);
            this.log(`         trust the Self-Signed Certificate for the site if you haven't done that)`);
            this.log(`      2. Sideload the add-in into your Office application.\n`);
            this.log(`      Please refer to resource.html in your project for more information.`);
            this.log(`      Or visit Office Add-ins repo at: https://github.com/officeDev/generator-office \n`);
            this.log('----------------------------------------------------------------------------------------------------------\n');
            this._exitProcess();
  }
  _detailedHelp() {
    this.log(`\nYo ${chalk.underline.bold.greenBright('Office-X')} ${chalk.bgGreen('Arguments' )} and ${chalk.bgMagenta('Options.')}\n`);
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
    this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`
    );
    this._exitProcess();
  }

  _exitProcess() {
    process.exit();
  }
};
