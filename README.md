# UniExtract
Universal Extractor 2 or Universal Extractor (Bioruebe version) as it is also called to avoid confusion is an unofficial updated and extended version of the [original UniExtract by Jared Breland](http://legroom.net/software/uniextract). As the development of the original version has stopped and no update has been published for years, many forked (modified and maintained by the community) Universal Extractors have arised. This is the most advanced and the only one with a fully transparent development using a public GitHub-powered source code repository.


## New features
- Batch mode
- Silent mode, not showing any prompts
- Scan only mode to determine file types without extracting
- Integrated updater
- 100+ new supported file types
- Audio and video extraction for multimedia files
- Cascading context menu
- Support for password list for common archives
- Improved optional status box with progress indicator
- New detection methods + more detailed output and error messages
- Support for some extractors not shipping with UniExtract as plugins
- Resource usage/speed improvements, lots of fixes
- Auto-using 64 bit versions of extractors if supported by OS

See the changelog for a complete log of all improvements.


## Portable version
Universal Extractor itself is completely portable, with some exceptions:
- Enabling context menu entries will create registry entries
- To extract a wide variety of file types more than 50 different extractors are used. Some of them might leave traces on the system. For the most common archives and installers extraction can be considered portable, for others probably not.
- Storing Universal Extractor in a directory without write access (e.g. C:\Program Files) enables multi-user mode. This results in configuration files being stored in the %APPDATA% directory (C:\Users\YourUsername\AppData\Roaming\Bioruebe\UniExtract).
See issue #20 for more information.


## Contributions
Any contribution in form of ideas, bug reports, code commits, documentation improvements, etc. is welcome. Help is currently needed in updating the translations for many languages. If you are able to translate into another language, take a look at the corresponding issue (#2) or open the language file in the `/lang` subdirectory and check for empty strings. English and German language are always up-to-date and can be used as a reference.

Feel free to submit bug reports or feature requests using the issues tab or the built-in feedback window in Universal Extractor, accessible via the 'Help' menu. See todo.txt for a list of notes on possible future changes.


## License
Universal Extractor is licensed under GPLv2. See LICENSE for the full legal text.
Code (functions, UDFs, etc.) written from scratch by me (which are not under copyleft) can also be used in your own projects under the terms of a BSD 3-clause license.

Universal Extractor uses [TrIDLib by Marco Pontello](http://mark0.net/code-tridlib-e.html) and many other great tools and libraries to support as many file formats as possible. A detailed list can be found [here](https://github.com/Bioruebe/UniExtract2/blob/master/helper_binaries_info.txt).