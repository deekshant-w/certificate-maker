# Certificate Maker
#### You make it, it bakes it...

This is simple project as the name suggests, "IT MAKES CERTIFICATES", not the cryptogrqaphy ones and can also email them if you want. Just provide it with a word template, a database to populate your certificates with, optional email credentials if you want. This software uses all that to create and email certificates. 

## Features

- Multi project system (Manage several template projects at a time)
- Excel database for ease of access
- Simple and clean GUI
- Easy to setup email service
- Free to use and extremely useful in several situations

## Installation

- Clone the repo
- Setup you virtualenv using the requirements.txt file
- Then just go for it...

> The software is not compiled but most the libraries used in it are "Windows" suitable only!

The `pdf` creation system and emailing system are entirely separated individually and can be easily altered for other operating systems. Ex. - `Gmail API` (too complex), `twisted` (unnecessarily complex) for mailing are good examples of cross platform mailing libraries. `docx2pdf` is excellent for batch processing for cross platform implementation.

## Email Feature
This feature requires the presence of a `creds.txt` file. It can be created by using the email feature once in any project but requires unsafe app password and account access. Do not share that file with anyone or store it in a repo. The retrieval of this password can be done using [this](https://support.google.com/accounts/answer/185833)



## Development

Want to contribute? Great! Do it...


## License

GNU General Public License v3.0

___
**Free Software, Hell Yeah!**

_I had to create a software for such a simple purpose because the alternatives are either paid or really terrible so use it change it, do whatever, just keep it open source and let me know so that I cann use ur software too!_
