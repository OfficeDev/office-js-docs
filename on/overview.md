# OneNote add-ins JavaScript API reference

_Applies to: OneNote Online_

The links below show the high level OneNote objects available in the APIs. Each object page link contains a description of the properties, relationships, and methods available on the object. Explore the links below to learn more.
	
* [OneNoteOM](resources/onenoteom.md): Represents the top level object which contains any globaly addressable objects such as Notebooks, Current Notebook, Current Section, etc. 
	* [Notebooks](resources/notebooks.md): Represents a collection of OneNote Notebooks.
* [Notebook ](resources/notebook.md): Represents the Notebook in OneNote. Contains SectionGroups and Sections as a child. 
	* [SectionGroups](resources/sectiongroups.md): Represents the SectionGroup in OneNote.  
	* [Sections](resources/sections.md): Represents a collection of Sections.  
* [SectionGroup](resources/sectiongroup.md): Represents the SectionGroup in OneNote.  Contains Sections and SectionGroups as child.
	* [SectionGroups](resources/sectiongroups.md): Represents the SectionGroup in OneNote.  
	* [Sections](resources/sections.md): Represents a collection of Sections.  
* [Section](resources/section.md): Represents the Section in OneNote.  Contains Pages as child.
	* [Pages](resources/pages.md): Represents a collection of Pages.  
* [Page](resources/page.md): Represents the Page in OneNote.  Contains PageContents as child.
	* [PageContents](resources/pagecontents.md): Represents a collection of PageContents.  
* [PageContent](resources/pagecontent.md): Represents a placeholder object for top leve content under page. These contents can be asigned x,y position in a page. It can be either an image or an outline.  
* [Outline](resources/outline.md): Represents a region that contains text. Contains paragraphs as child.  
	* [Paragraphs](resources/paragraphs.md): Collection of paragraph. 
* [Paragraph](resources/paragraph.md): Represents a placeholder object for contents in an Outline. It can either be rich text or image. Paragraph is auto laid out.  
* [Image](resources/image.md): Represents the Image object.  
* [RichText](resources/richtext.md): Represents RichText object. 


##### Additional resources

*  [Snippet Explorer for OneNote](http://officesnippetexplorer.azurewebsites.net/#/snippets/onenote)
*  [OneNote add-ins code samples](onenote-add-ins-code-samples.md) 


