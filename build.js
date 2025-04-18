// ESBuild configuration
import esbuild from 'esbuild';
import { nodeExternalsPlugin } from 'esbuild-node-externals';

// Build options
const commonOptions = {
  entryPoints: [ './index.ts' ],
  bundle: true,
  platform: 'neutral',
  target: [ 'es2018' ],
  sourcemap: false,
  minify: true,
  plugins: [ nodeExternalsPlugin() ],
};

// Create the main UMD build
async function build() {
  try {
    // Build the minified UMD bundle
    await esbuild.build({
      ...commonOptions,
      outfile: './build/parser.min.js',
      format: 'iife',
      globalName: 'FormulaParser',
    });

    // Build the ES module
    await esbuild.build({
      ...commonOptions,
      outfile: './build/parser.esm.js',
      format: 'esm',
    });

    // Build the CommonJS module
    await esbuild.build({
      ...commonOptions,
      outfile: './build/parser.cjs.js',
      format: 'cjs',
    });

    console.log('✅ Build completed successfully!');
  } catch (error) {
    console.error('❌ Build failed:', error);
    process.exit(1);
  }
}

// Run the build
build();
