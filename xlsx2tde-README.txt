README for xlsx2tde.py version 1.1

Before you start:

I recommend starting with the Anaconda distribution of Python, as it has many
of the tools needs in one package.

python modules required
    tableausdk
    tableau_rest_api
    argparse
    configparser

The above 4 do not come with most standard python installs, so they will need to
be added with something like pip. The tableausdk needs to be downloaded directlty
from Tableau's website. What I have done to make it easier, is included in the
zip file is a yml file that is a clone of the environment that I developed the
script in. Go to http://conda.pydata.org/docs/using/envs.html for more info
about how to use the file.

How it works:

This script takes a input Excel file and creates a tableau data extract (.tde)
that is uploaded to the site of the user's choosing. To do this, the sheets on
the file are condensed into on data frame with a module called pandas. Pandas
then turns the data frame into a csv. This is where any obvious clean up happens,
and the columns are organized in a more logical manner.

Once the data is a csv, a tde is generated. This is done line by line. Tableau is
able to process about 7500 rows/sec, so keep that in mind when you are trying to
determine how much time it will take.

To create the tde file, the script looks for a schema.ini file for instructions
on how to build the extract. This is important because it tells Tableau what data
type to use on each of the columns in the extract. Make sure to delete the old
version of the tde file before running the next iteration or it will simply append
the one previously created.

You can refer to the log files (Publish and DataExtract) to get more details on
what actually being done by the script.

The correct syntax when entering into the console is as follows:

python xlsx2tde.py -f XLSX_NAME -s SITE_NAME

The number of tabs in the excel file does not matter, but what is important is
the structure of the data. It will need to be 5 columns wide with this format:

Date
Site Name
Publisher Name
Campaign
Cost

After the tde is published to the site, the workbook that you would like to have
use it will need to be directed to it from Tableau Desktop. This only needs to be
done once. Subsequent uploads will simply replace that data on the server, so a
new connection will not need to be established.