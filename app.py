import os
import shutil
import pandas as pd
from pathlib import Path
from typing import List, Optional
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class FileHandler:
    """A class to handle file and folder operations for Excel workbook data updates."""
    
    def __init__(self, base_folder: str):
        """
        Initialize the FileHandler with a base folder.
        
        Args:
            base_folder (str): Path to the base folder
        """
        self.base_folder = Path(base_folder)
        self.create_folder(self.base_folder)
    
    def create_folder(self, folder_path: str) -> bool:
        """
        Create a folder if it doesn't exist.
        
        Args:
            folder_path (str): Path of the folder to create
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            Path(folder_path).mkdir(parents=True, exist_ok=True)
            logger.info(f"Folder created/verified: {folder_path}")
            return True
        except Exception as e:
            logger.error(f"Error creating folder {folder_path}: {e}")
            return False
    
    def create_project_structure(self, folders: List[str]) -> bool:
        """
        Create multiple folders inside the base folder.
        
        Args:
            folders (List[str]): List of folder names to create
            
        Returns:
            bool: True if all folders created successfully
        """
        success = True
        for folder in folders:
            folder_path = self.base_folder / folder
            if not self.create_folder(folder_path):
                success = False
        return success
    
    def copy_file(self, source_path: str, destination_path: str) -> bool:
        """
        Copy a file from source to destination.
        
        Args:
            source_path (str): Source file path
            destination_path (str): Destination file path
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Ensure destination directory exists
            dest_dir = Path(destination_path).parent
            self.create_folder(dest_dir)
            
            shutil.copy2(source_path, destination_path)
            logger.info(f"File copied: {source_path} -> {destination_path}")
            return True
        except Exception as e:
            logger.error(f"Error copying file {source_path} to {destination_path}: {e}")
            return False
    
    def move_file(self, source_path: str, destination_path: str) -> bool:
        """
        Move a file from source to destination.
        
        Args:
            source_path (str): Source file path
            destination_path (str): Destination file path
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Ensure destination directory exists
            dest_dir = Path(destination_path).parent
            self.create_folder(dest_dir)
            
            shutil.move(source_path, destination_path)
            logger.info(f"File moved: {source_path} -> {destination_path}")
            return True
        except Exception as e:
            logger.error(f"Error moving file {source_path} to {destination_path}: {e}")
            return False
    
    def copy_folder(self, source_folder: str, destination_folder: str) -> bool:
        """
        Copy an entire folder and its contents.
        
        Args:
            source_folder (str): Source folder path
            destination_folder (str): Destination folder path
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            shutil.copytree(source_folder, destination_folder, dirs_exist_ok=True)
            logger.info(f"Folder copied: {source_folder} -> {destination_folder}")
            return True
        except Exception as e:
            logger.error(f"Error copying folder {source_folder} to {destination_folder}: {e}")
            return False
    
    def move_folder(self, source_folder: str, destination_folder: str) -> bool:
        """
        Move an entire folder and its contents.
        
        Args:
            source_folder (str): Source folder path
            destination_folder (str): Destination folder path
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            shutil.move(source_folder, destination_folder)
            logger.info(f"Folder moved: {source_folder} -> {destination_folder}")
            return True
        except Exception as e:
            logger.error(f"Error moving folder {source_folder} to {destination_folder}: {e}")
            return False
    
    def list_files(self, folder_path: str, extension: Optional[str] = None) -> List[str]:
        """
        List all files in a folder, optionally filtered by extension.
        
        Args:
            folder_path (str): Path to the folder
            extension (str, optional): File extension to filter (e.g., '.xlsx')
            
        Returns:
            List[str]: List of file paths
        """
        try:
            folder = Path(folder_path)
            if extension:
                files = list(folder.glob(f"*{extension}"))
            else:
                files = [f for f in folder.iterdir() if f.is_file()]
            
            file_paths = [str(f) for f in files]
            logger.info(f"Found {len(file_paths)} files in {folder_path}")
            return file_paths
        except Exception as e:
            logger.error(f"Error listing files in {folder_path}: {e}")
            return []
    
    def delete_file(self, file_path: str) -> bool:
        """
        Delete a file.
        
        Args:
            file_path (str): Path to the file to delete
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            Path(file_path).unlink()
            logger.info(f"File deleted: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Error deleting file {file_path}: {e}")
            return False
    
    def delete_folder(self, folder_path: str) -> bool:
        """
        Delete a folder and all its contents.
        
        Args:
            folder_path (str): Path to the folder to delete
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            shutil.rmtree(folder_path)
            logger.info(f"Folder deleted: {folder_path}")
            return True
        except Exception as e:
            logger.error(f"Error deleting folder {folder_path}: {e}")
            return False
    
    def backup_excel_file(self, excel_path: str, backup_folder: str = None) -> str:
        """
        Create a backup of an Excel file with timestamp.
        
        Args:
            excel_path (str): Path to the Excel file
            backup_folder (str, optional): Backup folder path
            
        Returns:
            str: Path to the backup file
        """
        try:
            from datetime import datetime
            
            excel_file = Path(excel_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if backup_folder is None:
                backup_folder = self.base_folder / "backups"
            
            self.create_folder(backup_folder)
            
            backup_name = f"{excel_file.stem}_backup_{timestamp}{excel_file.suffix}"
            backup_path = Path(backup_folder) / backup_name
            
            self.copy_file(excel_path, str(backup_path))
            return str(backup_path)
        except Exception as e:
            logger.error(f"Error creating backup for {excel_path}: {e}")
            return ""
    
    def organize_excel_files(self, source_folder: str, organize_by: str = "date") -> bool:
        """
        Organize Excel files into subfolders based on criteria.
        
        Args:
            source_folder (str): Source folder containing Excel files
            organize_by (str): Organization criteria ('date', 'size', 'name')
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            excel_files = self.list_files(source_folder, '.xlsx')
            excel_files.extend(self.list_files(source_folder, '.xls'))
            
            for file_path in excel_files:
                file_obj = Path(file_path)
                
                if organize_by == "date":
                    # Organize by modification date
                    mod_time = file_obj.stat().st_mtime
                    from datetime import datetime
                    date_folder = datetime.fromtimestamp(mod_time).strftime("%Y-%m")
                    dest_folder = Path(source_folder) / "organized_by_date" / date_folder
                
                elif organize_by == "size":
                    # Organize by file size
                    size = file_obj.stat().st_size
                    if size < 1024 * 1024:  # < 1MB
                        size_folder = "small"
                    elif size < 10 * 1024 * 1024:  # < 10MB
                        size_folder = "medium"
                    else:
                        size_folder = "large"
                    dest_folder = Path(source_folder) / "organized_by_size" / size_folder
                
                else:  # organize by name (first letter)
                    first_letter = file_obj.stem[0].upper()
                    dest_folder = Path(source_folder) / "organized_by_name" / first_letter
                
                self.create_folder(dest_folder)
                dest_path = dest_folder / file_obj.name
                self.move_file(file_path, str(dest_path))
            
            return True
        except Exception as e:
            logger.error(f"Error organizing Excel files: {e}")
            return False


# Example usage and main function
def main():
    """Example usage of the FileHandler class."""
    
    # Initialize file handler with base folder
    handler = FileHandler("excel_project")
    
    # Create project structure
    project_folders = [
        "data/raw",
        "data/processed",
        "data/backup",
        "reports",
        "templates",
        "logs"
    ]
    
    print("Creating project structure...")
    handler.create_project_structure(project_folders)
    
    # Example operations
    print("\nExample operations:")
    
    # List files in a directory
    files = handler.list_files(".", ".py")
    print(f"Python files in current directory: {len(files)}")
    
    # Create a sample Excel file path (for demonstration)
    sample_excel = "sample_data.xlsx"
    
    # If you have an actual Excel file, you can:
    # backup_path = handler.backup_excel_file(sample_excel)
    # print(f"Backup created at: {backup_path}")
    
    print("File handler operations completed!")


if __name__ == "__main__":
    main()