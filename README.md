# Data-Whisperer
A NLP visualization builder for Tableau Desktop.

Requires Python 2.7 and the xml and tkFileDialog packages as well as the Personal Edition of Tableau Desktop for Excel connections or the Professional Edition for Excel and SQL connections.

To install Python 2.7 visit https://www.python.org/downloads/

For information on adding the required packages see: https://packaging.python.org/installing/

The .exe version of Data Whisperer was created using http://www.pyinstaller.org/

### Common Issues:

1) Some users have encountered issues with the automatic save location setting. To fix this simply enter a desired file location rather than hitting enter for the default Desktop location.

2) You will need to reconnect the sample workbook to the sample data source after downloading it.

3) Data Whisperer is only tested on Excel and SQL sources. Other sources may be functional but this is not guarenteed.

4) Data Whisperer is designed to be used with only workbooks that have a connection to a single data source.

5) The current query system does not handle more complex visualizations like maps, scatter plots, or creating calculations though these features are planned for future releases.
