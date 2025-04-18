import * as fs from 'fs';
import {createSyntaxDiagramsCode} from 'chevrotain';
import {Parser} from './parsing';

/**
 * Generate syntax diagrams for documentation
 */
function generateDiagrams() {
  // Create a parser instance
  const parserInstance = new Parser();

  // Extract the serialized grammar
  const serializedGrammar = parserInstance.getSerializedGastProductions();

  // Create the HTML text
  const htmlText = createSyntaxDiagramsCode(serializedGrammar);

  // Write the HTML file to disk
  fs.writeFileSync('./docs/generated_diagrams.html', htmlText);

  console.log('Generated syntax diagrams in ./docs/generated_diagrams.html');
}

// Execute if run directly
generateDiagrams();
