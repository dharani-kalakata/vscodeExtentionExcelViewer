const esbuild = require('esbuild');
const path = require('path');

async function main() {
  const isWatch = process.argv.includes('--watch');
  const isProduction = process.argv.includes('--production') || process.env.NODE_ENV === 'production';
  const outdir = path.join(__dirname, '..', 'media');

  const buildOptions = {
    entryPoints: [path.join(__dirname, '..', 'webview-src', 'main.jsx')],
    bundle: true,
    outdir,
    entryNames: 'webview',
    assetNames: 'assets/[name]',
    format: 'iife',
    sourcemap: isProduction ? false : 'inline',
    target: ['es2020'],
    loader: {
      '.css': 'css',
      '.png': 'file',
      '.svg': 'file',
      '.ttf': 'file',
      '.woff': 'file',
      '.woff2': 'file',
    },
    define: {
      'process.env.NODE_ENV': JSON.stringify(isProduction ? 'production' : 'development'),
    },
  };

  if (isWatch) {
    const ctx = await esbuild.context(buildOptions);
    await ctx.watch();
    console.log('Watching webview sources...');
  } else {
    await esbuild.build(buildOptions);
    console.log('Built webview assets.');
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
