"""
CalCareers Application Packager - Package Generator
Creates standardized folder structure and organizes documents.
"""

import os
from pathlib import Path
from typing import Dict, List
from core.data_models import PackageInput


class PackageGenerator:
    """Generates standardized CalCareers application packages."""
    
    # Standard folder structure
    FOLDER_STRUCTURE = [
        "01_Checklists",
        "02_STD678",
        "03_Required_Docs",
        "04_Optional_Docs",
        "05_VeteransPreference",
        "06_Receipts",
        "07_AuditLog"
    ]
    
    def __init__(self, output_root: str = "./outputs"):
        """
        Initialize package generator.
        
        Args:
            output_root: Base directory for all generated packages
        """
        self.output_root = Path(output_root)
        self.output_root.mkdir(exist_ok=True)
    
    def create_package_structure(self, pkg_input: PackageInput) -> Path:
        """
        Create standardized folder structure for application package.
        
        Args:
            pkg_input: Complete package input
            
        Returns:
            Path to created package root directory
        """
        package_name = pkg_input.get_package_name()
        package_path = self.output_root / package_name
        
        # Create main package directory
        package_path.mkdir(exist_ok=True)
        
        # Create standard subfolder structure
        for folder in self.FOLDER_STRUCTURE:
            folder_path = package_path / folder
            folder_path.mkdir(exist_ok=True)
            
            # Skip VeteransPreference folder if not claiming VP
            if folder == "05_VeteransPreference" and not pkg_input.veterans_preference.claiming_veterans_preference:
                # Remove the folder if created
                folder_path.rmdir()
        
        return package_path
    
    def generate_filename(self, pkg_input: PackageInput, doc_type: str, extension: str = "pdf") -> str:
        """
        Generate standardized filename.
        
        Args:
            pkg_input: Package input for naming
            doc_type: Document type (STD678, SOQ, Resume, etc.)
            extension: File extension without dot
            
        Returns:
            Standardized filename
        """
        prefix = pkg_input.get_filename_prefix()
        return f"{prefix}_{doc_type}.{extension}"
    
    def get_document_locations(self, pkg_input: PackageInput, decisions: Dict) -> Dict[str, str]:
        """
        Map document types to their target folders.
        
        Args:
            pkg_input: Package input
            decisions: Decisions dict from DecisionEngine
            
        Returns:
            Dict mapping doc_type to folder name
        """
        locations = {
            "STD678": "02_STD678",
            "Resume": "04_Optional_Docs",
            "SOQ": "04_Optional_Docs",
            "Transcript": "04_Optional_Docs",
            "Certs": "04_Optional_Docs"
        }
        
        # Veterans' Preference documents
        if pkg_input.veterans_preference.claiming_veterans_preference:
            vp_docs = {
                "DD214": "05_VeteransPreference",
                "CalHR1093": "05_VeteransPreference",
                "CalHR1094": "05_VeteransPreference",
                "MarriageCertificate": "05_VeteransPreference",
                "DeathCertificate": "05_VeteransPreference",
                "DisabilityAwardLetter": "05_VeteransPreference"
            }
            locations.update(vp_docs)
        
        return locations
    
    def create_placeholder_files(self, package_path: Path, pkg_input: PackageInput, 
                                 decisions: Dict) -> List[str]:
        """
        Create placeholder README files in each folder explaining what goes there.
        Does NOT create fake documents - only instructional placeholders.
        
        Args:
            package_path: Root package directory
            pkg_input: Package input
            decisions: Decisions from engine
            
        Returns:
            List of created placeholder paths
        """
        created = []
        prefix = pkg_input.get_filename_prefix()
        
        # Checklists folder
        checklist_readme = package_path / "01_Checklists" / "README.txt"
        checklist_readme.write_text(
            "This folder contains the VSC checklist and any application-specific notes.\n"
            f"Expected file: {prefix}_Checklist.txt\n"
        )
        created.append(str(checklist_readme))
        
        # STD 678 folder
        std678_readme = package_path / "02_STD678" / "README.txt"
        std678_readme.write_text(
            "This folder contains the completed STD 678 Application form.\n"
            f"Expected file: {prefix}_STD678.pdf\n"
            f"Template track: {decisions['template_track']}\n"
        )
        created.append(str(std678_readme))
        
        # Required docs folder
        required_readme = package_path / "03_Required_Docs" / "README.txt"
        required_list = "\n".join([f"- {doc}" for doc in decisions['base_required_docs']])
        required_readme.write_text(
            "This folder contains required documents per the job posting.\n"
            "Required documents:\n"
            f"{required_list}\n"
        )
        created.append(str(required_readme))
        
        # Optional docs folder
        optional_readme = package_path / "04_Optional_Docs" / "README.txt"
        optional_list = "\n".join([f"- {doc}" for doc in decisions.get('optional_docs', [])])
        optional_readme.write_text(
            "This folder contains optional supporting documents.\n"
            "Common optional documents:\n"
            f"{optional_list}\n"
        )
        created.append(str(optional_readme))
        
        # Veterans' Preference folder (if applicable)
        if pkg_input.veterans_preference.claiming_veterans_preference:
            vp_readme = package_path / "05_VeteransPreference" / "README.txt"
            vp_attachments = "\n".join([f"- {att}" for att in decisions['required_attachments']])
            vp_sections = ", ".join([str(s) for s in decisions['required_sections']])
            vp_readme.write_text(
                "This folder contains Veterans' Preference documentation.\n"
                f"Basis: {decisions['vp_basis']}\n"
                f"Required CalHR 1093 sections: {vp_sections}\n"
                "Required attachments:\n"
                f"{vp_attachments}\n"
                f"Complete CalHR 1093 form: {prefix}_CalHR1093.pdf\n"
            )
            created.append(str(vp_readme))
        
        # Receipts folder
        receipts_readme = package_path / "06_Receipts" / "README.txt"
        receipts_readme.write_text(
            "This folder contains submission receipts and confirmation emails.\n"
            f"Save CalCareers confirmation email here after submission.\n"
        )
        created.append(str(receipts_readme))
        
        # Audit log folder
        audit_readme = package_path / "07_AuditLog" / "README.txt"
        audit_readme.write_text(
            "This folder contains the automated audit log.\n"
            f"Expected file: {prefix}_AuditLog.json\n"
        )
        created.append(str(audit_readme))
        
        return created
