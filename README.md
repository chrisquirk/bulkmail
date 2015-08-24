# bulkmail

Tool for sending bulk mail using Exchange / Office365

You need to specify two things: a text template for the mail, and a list of recipients as a TSV file.
For each entry in the recipient list, an email will be sent.

## file formats

### maillist

The mailing list should be a TSV file.
The first line should be a header providing names for each of the columns; each header name should be unique.
Each subsequent line should have the same number of fields as the header line.
These values will be substituted into the template to produce a well formed message, then sent.

### template

The template file should consist of two parts.
The first part is a series of headers, one per line, ending with a blank line.
The remainder of the file is the HTML body of the message.

#### headers

The following headers may currently be used:

| header name | contents |
| -- | -- |
| subject | The subject of the mail. Should occur once. The last occurrence will overwrite any prior occurrences. |
| to | One recipient of the mail. This field can occur many timmes. Each occurrence should contain exactly one recipient email. |
| cc | Same as above, but for CC instead of To |
| bcc | Same as above, but for BCC instead of To |
| attach | Filename to be attached to the email |

Providing any other header will cause an error.

#### variables

Inside the template, variables may be used anywhere.
They will be replaced by the value provided in the mailing list before a message is created in sent.
Variable names surrounded by two square brackets will be substituted by their value.

## example

