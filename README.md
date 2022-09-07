# Flask HTML to PDF App

This app uses the Python Library Version of the Foxit PDF SDK available at the [Foxit Developers Download page](https://developers.foxit.com/download/). The Trial SN and KEY are found in the ZIP file.

You will also need the HTML for PDF Engine. The only way to get this is to contact [Foxit Support Center](https://kb.foxitsoftware.com/hc/en-us) and make the request. Once you get it unzip the package to the `engine` folder. The app won't work without this piece, unfortunately.

Install the following necessary packages with `pip`:

```bash
pip install Flask
pip install FoxitPDFSDKPython3
pip install cryptography
pip install pyOpenSSL
pip install pywin32
pip install uuid
pip install pandas
```

Run app with the following command:

```bash
flask --debug run
```
