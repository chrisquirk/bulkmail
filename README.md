# bulkmail

Tool for sending bulk mail using Exchange / Office365

You need to specify two things: a text template for the mail, and a list of recipients as a TSV file.  For each entry in the recipient list, an email will be sent.

The tool will prompt for your credentials every time.

## file formats

### maillist

The mailing list should be a TSV file.  The first line should be a header providing names for each of the columns; each header name should be unique.  Each subsequent line should have the same number of fields as the header line.  These values will be substituted into the template to produce a well formed message, then sent.

### template

The template file should consist of two parts.  The first part is a series of headers, one per line, ending with a blank line.  The remainder of the file is the HTML body of the message.

#### headers

The following headers may currently be used (note: header names are *not* case sensitive):

<dl>
<dt>subject</dt><dd>The subject of the mail. Should occur once. The last occurrence will overwrite any prior occurrences.</dd>
<dt>to</dt><dd>One recipient of the mail. This field can occur many times. Each occurrence should contain exactly one recipient email.</dd>
<dt>cc</dt><dd>Same as above, but for CC instead of To</dd>
<dt>bcc</dt><dd>Same as above, but for BCC instead of To</dd>
<dt>attach</dt><dd>Filename to be attached to the email</dd>
</dl>

Providing any other header will cause an error.

#### variables

Inside the template, variables may be used anywhere.  They will be replaced by the value provided in the mailing list before a message is created in sent.  Variable names surrounded by two square brackets will be substituted by their value. Note that variable names *are* case sensitive.

No error checking on variables is currentlly performed. If you mistype the name of a variable, it will not be replaced by anything, and no error will occur. Sorry.

## example

Consider the following example files:

### maillist.tsv

| Email | FirstName | LastName | Id |
| ----- | --------- | -------- | ---- |
| foo@bar.baz | Jane | Doe | 19 |
| abc@def.ghi | Edgar | Winters | 22 |

### template.txt

```html
Subject: test message to [[FirstName]] [[LastName]]
To: [[Email]]
Cc: cc_person_1@bar.bz
Cc: cc_person_2@bar.bz
Attach: file_[[Id]].txt

<p>Dear [[FirstName]],<p>

<p>Please see the attached file for more information.</p>

<p>This message was sent to [[Email]]</p>
```

This will result in two emails being sent. The first email will have attached the `file_19.txt` and look like this:
```html
Subject: test message to Jane Doe
To: foo@bar.baz
Cc: cc_person_1@bar.bz
Cc: cc_person_2@bar.bz
Attach: file_19.txt

<p>Dear Jane,<p>

<p>Please see the attached file for more information.</p>

<p>This message was sent to foo@bar.baz</p>
```

The second will have the attachment `file_22.txt` and have the corresponding replacements.
