Readme:
=======

NB!
This is preliminary work - do not look at this as a complete system!

If you see text in Liberum with exclamation marks (!) before and after they
indicate languagestring variables that doesn't exist in the current language

If you see text with at signs (@)before and after they indicate languagestrings
varibales that doesn't exist as variables.

You can force the language update to run (overriding the version check)
by calling setup.asp?force=1.

How to change text strings into variables:
==========================================

1. Create a new variable in tblLangStrings
2. Change the text to lang(cnnDB, "name-of-new-variable")
3. Add the variable and the english text to tblLangStrings
4. Add the translated text to other installed languages

NOTE - German Language:
======================
Due to differences in sentence structure between English and German,
admin/default.asp should be replaced with admin/default_german.asp.

ToDo:
=====
Admin form to export for a new translation for offline translations
  This form should create a new file with Update statements including all
  existing variables from tblLangString and empty translation fields for
  the new language.
Change all text into variables
Multi language options for statuses
Multi language options for e-mail messages


Translation Status
===========================
ALL PAGES TRANSLATED EXCEPT:
public.asp              <-- Done (Minus some error messages)
setup.asp               <-- Not translated
