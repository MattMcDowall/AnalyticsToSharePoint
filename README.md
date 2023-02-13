# AnalyticsToSharePoint

This is a fairly simple script to produce an Excel spreadsheet from an Alma Analytics report. Obviously you can do so manually, or use an Analytics Object in Alma to do so on a daily/weekly/monthly basis. But this script is to provide a bit more control, and allow the export both programmatically and on-demand.

## **Let me say up front . . .**
. . . that I’m relatively new and largely self-taught when it comes to both Git and Python. So I probably do things in not-the-most-efficient ways pretty regularly. If you have any advice for better methods to accomplish anything here, by all means please let me know.

## **Requirements**

- An Alma Analytics report, the specifics of which you can set within this script

### **Dependencies**
You’ll need Python to run the script, of course. You will also need the `pandas`, `requests` and `xmltodict` packages.

## **A note on authentication**
You’re also going to need appropriate API keys for your institution. Getting that set up is beyond the scope of this introduction, but I do want to mention how I’ve gone about implementing the authentication.

In order to keep the API credentials private, my approach is to tackle the authentication via an external file, and then call that file from these scripts.

What I’ve done is to create a script called `Credentials.py` which resides in this folder on my computer. To keep it private, I’ve added a line in the `.gitignore` file which filters it out of any pushes to the public repository. So you’ll need to create your own Credentials.py file, in this same directory.

Credentials.py is a very simple file, looking like this:

    prod_api = 'a1aa11111111111111aa1a11aaaa11a111aa'
    sand_api = 'z9zzz99zzz99z9z999zzzzz999zz9999z999'

As you can see in the scripts themselves, we import that file thus:

    import Credentials

and get the API key for use thus:

    apikey = Credentials.prod_api

I hope that makes sense.
