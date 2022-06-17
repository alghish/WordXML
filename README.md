# WordXML.Utilities.Engine
Library for processing word header , footer, content controller

Version 1.0.1

## Replacer Class

#### Header Image || Footer Image

```` C#
using WordXML;

string wordPath = "/example/test.docx";
string imageHeaderPath = "/example/header.png";
string imageFooterPath = "/example/footer.png";
string errorMsg = "";
		
Engine.Replacer.HeaderImage(wordPath, imageHeaderPath, ref errorMsg);
Engine.Replacer.FooterImage(wordPath, imageFooterPath, ref errorMsg)

````

#### Content Controller Copy

``` C#
using WordXML;

string wordSourcePath = "/example/source.docx";
string wordDestinationPath = "/example/dest.docx";		
string errorMsg = "";

Engine.Replacer.CopyContentController(wordSourcePath, wordDestinationPath, ref errorMsg);
```

## Remover Class

``` C#
using WordXML;

string wordPath = "/example/test.docx";	
string errorMsg = "";
		
Engine.Remover.FooterImageRemover(wordPath, ref errorMsg);
```
