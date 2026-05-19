"""
CalCareers Application Packager - Audit Logger
Creates machine-readable audit trails for package generation.
"""

import json
from pathlib import Path
from datetime import datetime
from typing import Dict, List
from core.data_models import PackageInput, AuditRecord


class AuditLogger:
    """Generates and manages audit logs for package generation."""
    
    def __init__(self):
        pass
    
    def create_audit_record(self, pkg_input: PackageInput, decisions: Dict, 
                           missing_data: List[str], authorities: List[str]) -> AuditRecord:
        """
        Create audit record for package generation.
        
        Args:
            pkg_input: Package input data
            decisions: Decisions made by DecisionEngine
            missing_data: List of missing data fields
            authorities: Governing authorities cited
            
        Returns:
            AuditRecord object
        """
        # Sanitize inputs (no SSNs, no sensitive info in plain text)
        inputs_received = self._sanitize_inputs(pkg_input)
        
        return AuditRecord(
            timestamp=datetime.now(),
            package_name=pkg_input.get_package_name(),
            inputs_received=inputs_received,
            decisions_made=decisions,
            governing_authorities=authorities,
            missing_data=missing_data
        )
    
    def _sanitize_inputs(self, pkg_input: PackageInput) -> Dict:
        """
        Sanitize input data for audit log.
        Removes sensitive info, keeps only necessary metadata.
        """
        return {
            "candidate": {
                "legal_name": pkg_input.candidate.legal_name,
                "email": pkg_input.candidate.email,
                "phone": pkg_input.candidate.phone,
                "has_ecos_id": pkg_input.candidate.ecos_id is not None,
                "education_count": len(pkg_input.candidate.education_entries),
                "work_experience_count": len(pkg_input.candidate.work_experience_entries)
            },
            "job": {
                "jc_number": pkg_input.job.jc_number,
                "classification_title": pkg_input.job.classification_title,
                "department": pkg_input.job.department,
                "final_filing_date": pkg_input.job.final_filing_date,
                "required_documents_count": len(pkg_input.job.required_documents)
            },
            "template_track": pkg_input.template_track.value,
            "veterans_preference": {
                "claiming": pkg_input.veterans_preference.claiming_veterans_preference,
                "basis": pkg_input.veterans_preference.application_basis.value if pkg_input.veterans_preference.application_basis else None
            }
        }
    
    def save_audit_log(self, audit_record: AuditRecord, package_path: Path) -> Path:
        """
        Save audit record to package.
        
        Args:
            audit_record: AuditRecord to save
            package_path: Root package directory
            
        Returns:
            Path to saved audit log file
        """
        audit_folder = package_path / "07_AuditLog"
        audit_folder.mkdir(exist_ok=True)
        
        # Generate filename
        prefix = audit_record.package_name.replace("_CalCareersPackage", "")
        audit_file = audit_folder / f"{prefix}_AuditLog.json"
        
        # Write JSON
        with open(audit_file, 'w') as f:
            json.dump(audit_record.to_dict(), f, indent=2)
        
        return audit_file
    
    def generate_missing_data_report(self, missing_data: List[str], package_path: Path) -> Path:
        """
        Generate human-readable missing data report.
        
        Args:
            missing_data: List of missing data fields
            package_path: Root package directory
            
        Returns:
            Path to missing data report
        """
        if not missing_data:
            return None
        
        audit_folder = package_path / "07_AuditLog"
        report_file = audit_folder / "MISSING_DATA_REPORT.txt"
        
        report_content = f"""{'='*80}
MISSING DATA REPORT
{'='*80}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

The following required data fields were not provided during package generation.
These fields must be completed before submission.

MISSING FIELDS:
{chr(10).join([f'- {field}' for field in missing_data])}

{'='*80}
ACTION REQUIRED: Complete all missing fields before proceeding with application.
{'='*80}
"""
        
        report_file.write_text(report_content)
        return report_file
