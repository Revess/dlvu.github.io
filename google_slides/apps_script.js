function exportPDFsFromFolder(parent_folder_id='', slides_folder_id='', slides=[]) {
    var slides_folder = DriveApp.getFolderById(slides_folder_id);
    var folderName = slides_folder.getName();
    var presentations = slides_folder.getFilesByType(MimeType.GOOGLE_SLIDES);
  
    var root_folder = DriveApp.getFolderById(parent_folder_id);  
  
    if(!root_folder.getFoldersByName('PDFs').hasNext()){
      // Create a folder in Google Drive to store the exported SVGs
      var outputFolder = DriveApp.createFolder('PDFs');
      outputFolder.moveTo(root_folder)
    } else {
      var outputFolder = root_folder.getFoldersByName('PDFs').next();
    };
  
    while (presentations.hasNext()) {
      var presentationFile = presentations.next();
      if(slides.includes(presentationFile.getId())){
        var pdf_pres = presentationFile.getAs('application/pdf');
        var fileName = presentationFile.getName() + '.pdf';
        outputFolder.createFile(pdf_pres).setName(fileName);
      }
    }
  }
  