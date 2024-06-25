'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const gutil = require('gulp-util');
const fs = require('fs');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

var getJson = function (file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
};

let bumpVersionSubTask = build.subTask('bump-version-subtask', function (gulp, buildOptions, done) {

  const currentCommand = buildOptions.args._[0];

  const skipFunc = gulp.src('./config/package-solution.json').pipe(gutil.noop());

  if (typeof currentCommand != "string") {
    gutil.log("The current command is undefined, skip version bump");
    return skipFunc;
  }

  const commandName = currentCommand.toLocaleLowerCase();

  if (commandName != "bundle" && commandName != "bump-version") {
    gutil.log("The current command is not 'bundle' or 'bump-version', skip version bump");
    return skipFunc;
  }

  const bumpVersion = commandName == "bump-version" || buildOptions.args["ship"] === true;

  if (!bumpVersion) {
    gutil.log("The current command is not 'bump-version' or the --ship argument was not specified, skip version bump");
    return skipFunc;
  }

  const a = buildOptions.args;

  const skipMajorVersion = typeof a["major"] == "undefined" || a["major"] === false;
  const skipMinorVersion = !skipMajorVersion || typeof a["minor"] == "undefined" || a["minor"] === false;
  const skipPatchVersion = !skipMajorVersion || !skipMinorVersion || a["patch"] === false;

  if (skipMajorVersion && skipMinorVersion && skipPatchVersion) {
    gutil.log("skip version bump, because all specified arguments (major, minor, patch) are set to 'false'")
    return skipFunc;
  }

  const pkgSolutionJson = getJson('./config/package-solution.json');
  const currentVersionNumber = String(pkgSolutionJson.solution.version);
  let nextVersionNumber = currentVersionNumber.slice();
  let nextVersionSplitted = nextVersionNumber.split('.');
  gutil.log('Current version: ' + currentVersionNumber);

  if (!skipMajorVersion) {
    nextVersionSplitted[0] = parseInt(nextVersionSplitted[0]) + 1;
    nextVersionSplitted[1] = 0;
    nextVersionSplitted[2] = 0;
    nextVersionSplitted[3] = 0;
  }

  if (!skipMinorVersion) {
    nextVersionSplitted[1] = parseInt(nextVersionSplitted[1]) + 1;
    nextVersionSplitted[2] = 0;
    nextVersionSplitted[3] = 0;
  }

  if (!skipPatchVersion) {
    nextVersionSplitted[2] = parseInt(nextVersionSplitted[2]) + 1;
    nextVersionSplitted[3] = 0;
  }

  nextVersionNumber = nextVersionSplitted.join(".");

  gutil.log('New version: ', nextVersionNumber);

  pkgSolutionJson.solution.version = nextVersionNumber;
  fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolutionJson, null, 4), () => { });

  const packageJson = getJson('./package.json');
  packageJson.version = nextVersionNumber.split('.').splice(0, 3).join(".");
  fs.writeFile('./package.json', JSON.stringify(packageJson, null, 4), () => { });

  return gulp.src('./config/package-solution.json')
    .pipe(gulp.dest('./config'));
});

let bumpVersionTask = build.task('bump-version', bumpVersionSubTask);
build.rig.addPreBuildTask(bumpVersionTask);

build.initialize(gulp);
