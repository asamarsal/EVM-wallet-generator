import chalk from 'chalk';
import { Wallet, ethers } from 'ethers';
import { appendFileSync } from 'fs';
import moment from 'moment';
import readlineSync from 'readline-sync';
import xlsx from 'xlsx';
import json2csv from 'json2csv'; // Import json2csv properly

// Function to create a new Ethereum account
function createAccountETH() {
  const wallet = ethers.Wallet.createRandom();
  const privateKey = wallet.privateKey;
  const publicKey = wallet.publicKey;
  const mnemonicKey = wallet.mnemonic.phrase;

  return { privateKey, publicKey, mnemonicKey, address: wallet.address };
}

// Function to generate .xlsx file from wallet data
function generateXLSX(wallets) {
  const data = wallets.map((wallet, index) => ({
    No: index + 1,
    Address: wallet.address,
    PrivateKey: wallet.privateKey,
    Mnemonic: wallet.mnemonicKey,
  }));

  const ws = xlsx.utils.json_to_sheet(data, {
    header: ['No', 'Address', 'PrivateKey', 'Mnemonic'],  // Set custom header
  });

  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Wallets');

  // Centering the Address column by setting column widths and alignment
  const range = xlsx.utils.decode_range(ws['!ref']);
  for (let col = range.s.c; col <= range.e.c; col++) {
    const colLetter = xlsx.utils.encode_col(col);
    const maxLength = Math.max(
      ...Object.keys(ws)
        .filter((key) => key.startsWith(colLetter))
        .map((key) => ws[key]?.v ? ws[key].v.toString().length : 0)
    );
    ws['!cols'] = ws['!cols'] || [];
    ws['!cols'][col] = { width: maxLength + 2 }; // Adjust column width to fit data
  }

  // Set alignment for Address column to center
  for (let row = range.s.r; row <= range.e.r; row++) {
    const cell = ws[xlsx.utils.encode_cell({ r: row, c: 1 })];  // Address column
    if (cell) cell.s = { alignment: { horizontal: 'center' } };
  }

  xlsx.writeFile(wb, 'result.xlsx');
  console.log(chalk.green('Generated result.xlsx successfully.'));
}

// Function to generate .csv file from wallet data
function generateCSV(wallets) {
  const data = wallets.map((wallet, index) => ({
    No: index + 1,
    Address: wallet.address,
    PrivateKey: wallet.privateKey,
    Mnemonic: wallet.mnemonicKey,
  }));

  const csv = json2csv.parse(data); // Use json2csv to parse data into CSV format
  appendFileSync('result.csv', csv);
  console.log(chalk.green('Generated result.csv successfully.'));
}

// Main function using async IIFE (Immediately Invoked Function Expression)
(async () => {
  try {
    const wallets = [];
    // Get the total number of wallets to create from user input
    const totalWallet = readlineSync.question(
      chalk.yellow('Input how many wallets you want to create: ')
    );

    let count = 1;

    // If the user entered a valid number greater than 1, set the count
    if (totalWallet > 1) {
      count = totalWallet;
    }

    // Create the specified number of wallets
    while (count > 0) {
      const createWalletResult = createAccountETH();
      const theWallet = new Wallet(createWalletResult.privateKey);

      if (theWallet) {
        // Append wallet details to result.txt
        appendFileSync(
          './result.txt',
          `Address: ${theWallet.address} | Private Key: ${createWalletResult.privateKey} | Mnemonic: ${createWalletResult.mnemonicKey}\n`
        );
        appendFileSync('./address.txt', `${theWallet.address}\n`);
        appendFileSync('./privatekey.txt', `${createWalletResult.privateKey}\n`);
        
        // Display success message with the wallet address and timestamp
        console.log(
          chalk.green(
            `[${moment().format('HH:mm:ss')}] => ` + 'Wallet created...! Your address : ' + theWallet.address
          )
        );
        
        // Store wallet data for later generation of .xlsx or .csv
        wallets.push(createWalletResult);
      }

      count--;
    }

    // Display final message after creating all wallets
    console.log(
      chalk.green(
        'All wallets have been created. Check result.txt, address.txt, privatekey.txt for your results.'
      )
    );

    // Prompt user to choose whether to generate .xlsx, .csv or finish the program
    let choice = null;
    while (choice !== '0') {
      console.log(chalk.yellow('Do you want to generate .xlsx or .csv? Press 1 for .xlsx, 2 for .csv, and 0 for done : '));
      console.log('1. Generate .xlsx');
      console.log('2. Generate .csv');
      console.log('0. Done');
      choice = readlineSync.question('Your choice: ');

      if (choice === '1') {
        generateXLSX(wallets);
      } else if (choice === '2') {
        generateCSV(wallets);
      } else if (choice === '0') {
        console.log(chalk.green('Program ended.'));
      } else {
        console.log(chalk.red('Invalid choice! Please choose 1, 2, or 0.'));
      }
    }

    return;
  } catch (error) {
    // Display error message if an error occurs
    console.log(chalk.red('Your program encountered an error! Message: ' + error));
  }
})();
