# Command line AddressBook.framework-related tools, almost in MacRuby

I'm an oddball in that I use Microsoft Excel® for many things in life. Address books, for example. I have been keeping my address books using Excel since 1996. I have my own formats and rules, which also have evolved with time.

I switched to Mac in 2004 and I don't like its Address Book, just like I don't like many others.

The problem of most address book apps and formats is that it's designed by a person who has (almost) never travelled or never has many non-domestic contacts or even never visits a barber shop. Seriously. Many of them are good at strictly one setting (business or personal) but not at others. Many of them are good when used in one country but not in others. And we all know what happened to the ever-present Fax Number field. They become irrelevant. And as for MSN/AIM/Skype/Twitter/Facebook/etc.: If they already keep a list of your contacts on them, what's the point of keeping a separate record of them in your address book?

That's why I'm still sticking to my messy spreadsheets.

This being said, I need to bring them into OS X Address Book so that they sync with my iPhone (and previously many phones that sucked in one way or another; the only other phone that didn't suck was a Nokia 6150, but that was in my pre-Mac days and because 6150 never broke down Nokia probably found it too good to keep making phones like them, and there was no more any good, solid Series 40 phone like 6150).

Playing with AddressBook.framework is a pain but gladly we have [MacRuby](http://www.macruby.org/).

So this tool in its present form has two parts. The first part is a Perl script that converts my Excel spreadsheet to YAML, then a MacRuby script that reads the CSV into my Address Book.

Why Perl? Because the Ruby version of Spreadsheet::ParseExcel is not really usable. Try `xls_to_yaml.rb` with a sparse worksheet (i.e. worksheet with lots of empty cells in between the data cells) and you'll see. Plus I've tried that with MacRuby, and it's not yet working there. 

This being said, the MacRuby part has also its own oddities. String split doesn't always work, for example, hence the very awkward `-[NSString componentsSeparatedByString:]`.

To convert my Excel address book to Address Book, I use this pipe magic:

    ./xls_to_yaml.pl AddressBook.xls | ./yaml_to_ab.rb

My Excel address book has these columns:

*   tags: such as "work client" or "service"
*   nickname
*   name: in the form of "Last Name, First Name"
*   kanji: type a space between the last name and first name
*   phones: semicolon-separated, preferably forms like +1-123-456-789p1234
*   emails: semicolon-separated
*   addresses: semicolon-separated
*   title // org: such as "CEO // Acme, LLC"
*   notes

The Perl script depends on these CPAN modules:

*   `Jcode`
*   `Spreadsheet::ParseExcel`
*   `YAML`

And you'll need MacRuby 0.3 or above to run the ruby script

# Copying

These scripts are public domain.
