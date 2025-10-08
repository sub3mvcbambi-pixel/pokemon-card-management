#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const value = argv[i];
    if (value.startsWith('--')) {
      const [key, val] = value.replace(/^--/, '').split('=');
      args[key] = val ?? argv[i + 1];
      if (val === undefined) {
        i += 1;
      }
    } else if (!args._) {
      args._ = [];
      args._.push(value);
    } else {
      args._.push(value);
    }
  }
  return args;
}

function resolveScriptId(argvArgs) {
  if (argvArgs.scriptId) {
    return argvArgs.scriptId;
  }
  if (argvArgs._ && argvArgs._.length > 0) {
    return argvArgs._[0];
  }
  if (process.env.GAS_SCRIPT_ID) {
    return process.env.GAS_SCRIPT_ID;
  }
  if (process.env.SCRIPT_ID) {
    return process.env.SCRIPT_ID;
  }
  return null;
}

function writeClaspConfig(scriptId) {
  const targetPath = path.join(process.cwd(), '.clasp.json');
  const payload = {
    scriptId,
    rootDir: '.',
  };
  fs.writeFileSync(targetPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  return targetPath;
}

function main() {
  const argvArgs = parseArgs(process.argv.slice(2));
  const scriptId = resolveScriptId(argvArgs);

  if (!scriptId) {
    console.error('Usage: npm run link -- --scriptId <Apps Script ID>');
    console.error('You can also set GAS_SCRIPT_ID or SCRIPT_ID environment variables.');
    process.exitCode = 1;
    return;
  }

  const targetPath = writeClaspConfig(scriptId);
  console.log(`Created ${targetPath} targeting script ${scriptId}`);
  console.log('Run "npm run deploy" to push the code to the Apps Script project.');
}

if (require.main === module) {
  try {
    main();
  } catch (error) {
    console.error('Failed to create .clasp.json:', error.message);
    process.exitCode = 1;
  }
}