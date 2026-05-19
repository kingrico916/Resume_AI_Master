"""
CalCareers Application Packager - Main Orchestrator
Entry point for package generation automation.
"""

from pathlib import Path
from typing import Optional
from core.data_models import PackageInput
from core.decision_engine import DecisionEngine
from core.package_generator import PackageGenerator
from core.checklist_generator import ChecklistGenerator
from core.audit_logger import AuditLogger


class CalCareersPackager:
    """
    Main orchestrator for CalCareers Application Packager.
    Coordinates decision engine, package generation, and audit logging.
    """
    
    def __init__(self, output_root: str = "./outputs"):
        """
        Initialize packager.
        
        Args:
            output_root: Base directory for all generated packages
        """
        self.decision_engine = DecisionEngine()
        self.package_generator = PackageGenerator(output_root)
        self.checklist_generator = ChecklistGenerator()
        self.audit_logger = AuditLogger()
        
    def generate_package(self, pkg_input: PackageInput, validate: bool = True) -> dict:
        """
        Generate complete CalCareers application package.
        
        Args:
            pkg_input: Complete package input
            validate: Whether to validate inputs (default True)
            
        Returns:
            Dict with generation results and paths
        """
        results = {
            "success": False,
            "package_path": None,
            "checklist_path": None,
            "audit_log_path": None,
            "missing_data_report_path": None,
            "errors": [],
            "warnings": []
        }
        
        # Step 1: Validate inputs
        missing_data = []
        if validate:
            is_valid, missing_data = self.decision_engine.validate_inputs(pkg_input)
            if not is_valid:
                results["errors"].append(f"Validation failed. Missing {len(missing_data)} required fields.")
                results["warnings"].append("Package will be generated with missing data flagged.")
        
        # Step 2: Run decision engine
        try:
            decisions = self.decision_engine.analyze_package(pkg_input)
            authorities = self.decision_engine.get_authorities()
        except Exception as e:
            results["errors"].append(f"Decision engine error: {str(e)}")
            return results
        
        # Step 3: Create package structure
        try:
            package_path = self.package_generator.create_package_structure(pkg_input)
            results["package_path"] = str(package_path)
        except Exception as e:
            results["errors"].append(f"Package structure creation error: {str(e)}")
            return results
        
        # Step 4: Create placeholder files
        try:
            self.package_generator.create_placeholder_files(package_path, pkg_input, decisions)
        except Exception as e:
            results["warnings"].append(f"Placeholder creation warning: {str(e)}")
        
        # Step 5: Generate checklist
        try:
            checklist_content = self.checklist_generator.generate_checklist(pkg_input, decisions)
            checklist_path = package_path / "01_Checklists" / f"{pkg_input.get_filename_prefix()}_Checklist.txt"
            checklist_path.write_text(checklist_content, encoding='utf-8')
            results["checklist_path"] = str(checklist_path)
        except Exception as e:
            results["errors"].append(f"Checklist generation error: {str(e)}")
            return results
        
        # Step 6: Create audit log
        try:
            audit_record = self.audit_logger.create_audit_record(
                pkg_input, decisions, missing_data, authorities
            )
            audit_log_path = self.audit_logger.save_audit_log(audit_record, package_path)
            results["audit_log_path"] = str(audit_log_path)
        except Exception as e:
            results["warnings"].append(f"Audit log warning: {str(e)}")
        
        # Step 7: Generate missing data report (if applicable)
        if missing_data:
            try:
                missing_report_path = self.audit_logger.generate_missing_data_report(missing_data, package_path)
                results["missing_data_report_path"] = str(missing_report_path)
            except Exception as e:
                results["warnings"].append(f"Missing data report warning: {str(e)}")
        
        # Success
        results["success"] = True
        return results
    
    def print_results(self, results: dict) -> None:
        """
        Print generation results in a clean format.
        
        Args:
            results: Results dict from generate_package()
        """
        print("\n" + "="*80)
        print("CALCAREERS APPLICATION PACKAGER - GENERATION RESULTS")
        print("="*80)
        
        if results["success"]:
            print("\n✓ Package generated successfully")
            print(f"\nPackage location: {results['package_path']}")
            print(f"Checklist: {results['checklist_path']}")
            print(f"Audit log: {results['audit_log_path']}")
            if results["missing_data_report_path"]:
                print(f"Missing data report: {results['missing_data_report_path']}")
        else:
            print("\n✗ Package generation failed")
        
        if results["errors"]:
            print("\nERRORS:")
            for error in results["errors"]:
                print(f"  - {error}")
        
        if results["warnings"]:
            print("\nWARNINGS:")
            for warning in results["warnings"]:
                print(f"  - {warning}")
        
        print("\n" + "="*80 + "\n")


def main():
    """Example usage demonstration."""
    from core.data_models import (
        CandidateProfile, JobTarget, VeteransPreference,
        EducationEntry, WorkExperienceEntry,
        TemplateTrack, ApplicationBasis, PackageInput
    )
    
    # Example candidate profile
    candidate = CandidateProfile(
        legal_name="John Doe",
        date_of_birth="1990-01-15",
        address="123 Main St, Sacramento, CA 95814",
        phone="916-555-0100",
        email="john.doe@example.com",
        education_entries=[
            EducationEntry(
                institution="University of California",
                degree="Bachelor of Science",
                major="Computer Science",
                graduation_date="2012-05"
            )
        ],
        work_experience_entries=[
            WorkExperienceEntry(
                employer="State of California",
                job_title="IT Specialist",
                start_date="2015-06",
                end_date=None,
                duties="Network administration and system support"
            )
        ]
    )
    
    # Example job target
    job = JobTarget(
        jc_number="JC-123456",
        classification_title="Information Technology Specialist I",
        department="Department of Technology"
    )
    
    # Example veterans' preference
    vp = VeteransPreference(
        claiming_veterans_preference=True,
        application_basis=ApplicationBasis.SELF_VETERAN
    )
    
    # Create package input
    pkg_input = PackageInput(
        candidate=candidate,
        job=job,
        template_track=TemplateTrack.IT,
        veterans_preference=vp
    )
    
    # Generate package
    packager = CalCareersPackager()
    results = packager.generate_package(pkg_input)
    packager.print_results(results)


if __name__ == "__main__":
    main()
