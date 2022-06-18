using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordXML;

public static class Engine
{
	public static class Replacer
	{

	  public static string HeaderImage(string wordFilePath, string imagePath, ref string errorMessage) 
	  {
	  	try 
	  	{
		  	using (WordprocessingDocument document = WordprocessingDocument.Open(wordFilePath, true))
	        {
	            string headerRelationshipId;
	            string headerType = "";
	            //StringBuilder headerText = null;
	            Document mainDocument = document.MainDocumentPart.Document;

	            // Iterate through the headerReference elements in the main
	            // document part.
	            foreach (HeaderReference headerReference in
	                mainDocument.Descendants<HeaderReference>())
	            {
	                // The headerReference element has an explicit relationship
	                // with a Header part. Get the relationship id that points
	                // to the header part.
	                headerRelationshipId = headerReference.Id.Value;

	                // Get the header's type from the headerReference
	                // Type attribute.
	                headerType = headerReference.Type.Value.ToString();

	                // Get the header element from the Header part via the
	                // explicit relationship id.

	                HeaderPart headerPart = (HeaderPart)(document.MainDocumentPart.GetPartById(headerRelationshipId));

	                //    headerRelationshipId, headerType, headerText);
	                if (headerType == "Default")
	                {
	                    ImagePart imagePart = headerPart.ImageParts.First();
	                    byte[] imageBytes = File.ReadAllBytes(imagePath);
	                    BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
	                    writer.Write(imageBytes);
	                    writer.Close();
	                }
	            }
	        }
		  	
	  	} catch (System.Exception e) 
	  	{
	  		errorMessage = e.Message;
	  		return "Failed";
	  	} finally 
	  	{
	  		
	  	}
	  	
	  	return "Success";
	  }

	  public static string FooterImage(string wordFilePath, string imagePath, ref string errorMessage) 
	  {
	  	try 
	  	{
	  		using (WordprocessingDocument document = WordprocessingDocument.Open(wordFilePath, true))
	        {
	            string footerrelationshipid;
	            string footertype = null!;
	            //stringbuilder headertext = null;
	            Document maindocument = document.MainDocumentPart.Document;


	            // iterate through the headerreference elements in the main
	            // document part.
	            foreach (FooterReference footerreference in
	                maindocument.Descendants<FooterReference>())
	            {
	                // the headerreference element has an explicit relationship
	                // with a header part. get the relationship id that points
	                // to the header part.
	                footerrelationshipid = footerreference.Id.Value;

	                // get the header's type from the headerreference
	                // type attribute.
	                footertype = footerreference.Type?.Value.ToString();

	                // get the header element from the header part via the
	                // explicit relationship id.

	                FooterPart footerpart = (FooterPart)(document.MainDocumentPart.GetPartById(footerrelationshipid));

	                if (footertype != null && footerpart != null && footertype == "default")
	                {
	                    ImagePart imagepart = footerpart.ImageParts.First();
	                    byte[] imagebytes = File.ReadAllBytes(imagePath);
	                    BinaryWriter writer = new BinaryWriter(imagepart.GetStream());
	                    writer.Write(imagebytes);
	                    writer.Close();
	                }
	            }
	        }

	  	} catch (System.Exception e) 
	  	{
	  		errorMessage = e.Message;
	  		return "Failed";
	  	} finally 
	  	{
	  		
	  	}
	  	
	  	return "Success";
	  }

	  public static string CopyContentController(string sourceWordFilePath, string destinationWordFilePath, ref string errorMessage) 
	  {
	  	try 
	  	{
				if (File.Exists(destinationWordFilePath))
                {

                        // Start copy content of word files
                        string copy_txt = string.Empty;
                        string copy_xml = string.Empty;

                        using (WordprocessingDocument doc = WordprocessingDocument.Open(sourceWordFilePath, true))
                        {
                            MainDocumentPart mainPart = doc.MainDocumentPart;

                            List<SdtElement> controls = mainPart.Document.Body.Descendants<SdtElement>().Where(r =>
                            {
                                var tag = r.SdtProperties.GetFirstChild<Tag>();
                                return tag != null;
                            }).ToList();

                            foreach (var control in controls)
                            {
                                if (control != null)
                                {
                                    //SdtProperties props = control.Elements<SdtProperties>().FirstOrDefault();
                                    //Tag tag = props.Elements<Tag>().FirstOrDefault();
                                    //Console.WriteLine("Tag: " + tag.Val);


                                    copy_txt = control.InnerText;
                                    copy_xml = control.InnerXml;
                                }
                            }

                        }

                        using (WordprocessingDocument doc = WordprocessingDocument.Open(destinationWordFilePath, true))
                        {
                            MainDocumentPart mainPart = doc.MainDocumentPart;

                            List<SdtElement> controls = mainPart.Document.Body.Descendants<SdtElement>().Where(r =>
                            {
                                var tag = r.SdtProperties.GetFirstChild<Tag>();
                                return tag != null;
                            }).ToList();

                            foreach (var control in controls)
                            {
                                if (control != null)
                                {

                                    control.InnerXml = copy_xml;

                                }
                            }

                            mainPart.Document.Save();
                        }
                        //

                }


	  	} catch (Exception e) 
	  	{
	  		errorMessage = e.Message;
	  		return "Failed";
	  	} finally 
	  	{
	  		
	  	}
	  	return "Success";
	  }
	}

	public static class Remover 
	{
	  
	  public static string FooterImageRemover(string wordFilePath, ref string errorMessage) 
	  {
	  	
	  	try 
	  	{
	  		using (WordprocessingDocument document = WordprocessingDocument.Open(wordFilePath, true))
	        {
	            //StringBuilder headerText = null;
	            Document mainDocument = document.MainDocumentPart.Document;

	            // remove footer
	            if (document.MainDocumentPart.FooterParts.Count() > 0)
	            {
	                document.MainDocumentPart.DeleteParts(document.MainDocumentPart.FooterParts);
	            }
	            var footers =
	                  mainDocument.Descendants<FooterReference>().ToList();

	            foreach (var footer in footers)
	            {
	                footer.Remove();
	            }

	            //mainDocument.Save();
	            SectionProperties sectionProps = new SectionProperties();
	            PageMargin pageMargin = new PageMargin() { Top = 1, Right = (UInt32Value)1008U, Bottom = 1008, Left = (UInt32Value)1008U, Header = (UInt32Value)1U, Footer = (UInt32Value)1U, Gutter = (UInt32Value)0U };
	            sectionProps.Append(pageMargin);
	            document.MainDocumentPart.Document.Body.Append(sectionProps);
	        }

	  	} catch (System.Exception e) 
	  	{
	  		errorMessage = e.Message;
	  		return "Failed";
	  	} finally 
	  	{
	  		
	  	}

	  	return "Success";
	  }

	}
}
