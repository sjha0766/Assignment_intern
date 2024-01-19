const fs = require('fs');
const path = require('path');
const { program } = require('commander');
const filterData=require('./filter');

program
  .version('1.0.0')
  .description('A console application to save an Excel file locally and display its data.')
  .arguments('<file>')
  .action(file => {
    program.file = file;
  });

program.parse(process.argv);

if (!program.file) {
  console.error('Error: Please provide a valid file path.');
  process.exit(1);
}

function saveFileLocally(filePath) {
  const destinationPath = 'uploads/';

  // Create the 'uploads' directory if it doesn't exist
  if (!fs.existsSync(destinationPath)) {
    fs.mkdirSync(destinationPath);
  }

  const fileName = `uploaded_file_${Date.now()}${path.extname(filePath)}`;
  const destinationFile = path.join(destinationPath, fileName);

  // Copy the file to the 'uploads' directory
  fs.copyFileSync(filePath, destinationFile);

  console.log(`File saved locally at: ${destinationFile}`);

  return destinationFile;
}

function main() {
  const filePath = path.resolve(program.file);

  // Check if the file exists

  if (!fs.existsSync(filePath)) {
    console.error('Error: File does not exist. Please provide a valid file path.');
    process.exit(1);
  }

  console.log(`Saving file locally: ${filePath}`);
  const savedFilePath = saveFileLocally(filePath);

  filterData(savedFilePath);
 
}

main();
