"""
Final documentation generator module
Framework-compatible implementation with basic documentation synthesis.
"""

from typing import List, Optional, Dict, Any
import json
from pathlib import Path
from dataclasses import dataclass

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory

@dataclass
class DocumentationData:
    file_overview: Dict[str, Any]
    executive_summary: str
    detailed_analysis: Dict[str, Any]
    recommendations: List[str]
    ai_navigation_guide: Dict[str, Any]
    metadata: Dict[str, Any]


class DocSynthesizer(BaseAnalyzer):
    """Framework-compatible documentation synthesizer"""
    
    def __init__(self, name: str = "doc_synthesizer", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["structure_mapper", "data_profiler"])
    
    def _perform_analysis(self, context: AnalysisContext) -> DocumentationData:
        """Synthesize final documentation from all module results
        
        Args:
            context: AnalysisContext with access to all module results
            
        Returns:
            Dict with synthesized documentation
        """
        try:
            # Collect all available module results
            documentation = {
                'file_overview': self._create_file_overview(context),
                'executive_summary': self._create_executive_summary(context),
                'detailed_analysis': self._collect_detailed_analysis(context),
                'recommendations': self._generate_recommendations(context),
                'ai_navigation_guide': self._create_ai_navigation_guide(context),
                'metadata': self._create_metadata(context)
            }
            
            return DocumentationData(**documentation)
            
        except Exception as e:
            # Return minimal documentation rather than failing
            self.logger.error(f"Documentation synthesis failed: {e}")
            return DocumentationData(
                file_overview={'error': str(e)},
                executive_summary=f"Documentation synthesis failed: {e}",
                detailed_analysis={},
                recommendations=[],
                ai_navigation_guide={},
                metadata={'synthesis_error': str(e)}
            )
    
    def _validate_result(self, data: DocumentationData, context: AnalysisContext) -> ValidationResult:
        """Validate documentation synthesis results
        
        Args:
            data: Documentation data to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness based on documentation sections
        required_sections = ['file_overview', 'executive_summary', 'detailed_analysis', 'recommendations']
        present_sections = sum(1 for section in required_sections if getattr(data, section, None))
        completeness = present_sections / len(required_sections)
        
        # Accuracy based on successful module synthesis
        if 'error' in data.file_overview:
            accuracy = 0.3
            validation_notes.append("File overview generation failed")
        elif not data.executive_summary:
            accuracy = 0.5
            validation_notes.append("Executive summary missing")
        else:
            accuracy = 0.9
        
        # Consistency - check for logical structure
        consistency = 0.9
        if not data.metadata:
            consistency -= 0.2
            validation_notes.append("Metadata missing")
        
        # Confidence based on completeness and accuracy
        if completeness > 0.8 and accuracy > 0.7:
            confidence = ConfidenceLevel.HIGH
        elif completeness > 0.6 and accuracy > 0.5:
            confidence = ConfidenceLevel.MEDIUM
        else:
            confidence = ConfidenceLevel.LOW
        
        if len(data.recommendations) > 0:
            validation_notes.append("Recommendations generated")
        if data.ai_navigation_guide:
            validation_notes.append("AI navigation guide created")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _create_file_overview(self, context: AnalysisContext) -> Dict[str, Any]:
        """Create high-level file overview
        
        Args:
            context: AnalysisContext with module results
            
        Returns:
            Dict with file overview information
        """
        overview = {
            'file_name': context.file_path.name,
            'file_size_mb': context.file_metadata.file_size_mb,
            'analysis_timestamp': context.analysis_start_time,
            'processing_summary': {}
        }
        
        try:
            # Get basic file information from health checker
            health_result = context.get_module_result("health_checker")
            if health_result and health_result.data:
                overview['file_accessible'] = health_result.data.file_accessible
                overview['security_issues'] = len(health_result.data.security_issues)
                overview['excel_version'] = health_result.data.excel_version
            
            # Get structure information
            structure_result = context.get_module_result("structure_mapper")
            if structure_result and structure_result.data:
                overview['worksheet_count'] = structure_result.data.worksheet_count
                overview['has_hidden_sheets'] = len(structure_result.data.hidden_sheets) > 0
                overview['named_ranges_count'] = len(structure_result.data.named_ranges)
                overview['chart_count'] = structure_result.data.chart_count
                overview['pivot_table_count'] = structure_result.data.pivot_table_count
            
            # Get data quality information
            data_result = context.get_module_result("data_profiler")
            if data_result and data_result.data:
                overview['data_quality_score'] = data_result.data.data_quality_score
                overview['sheets_with_data'] = len([
                    sheet for sheet, profile in data_result.data.sheet_profiles.items()
                    if isinstance(profile, dict) and profile.get('status') == 'success'
                ])
            
            # Get complexity indicators
            formula_result = context.get_module_result("formula_analyzer")
            if formula_result and formula_result.data:
                overview['formula_count'] = formula_result.data.total_formulas
                overview['formula_complexity'] = formula_result.data.formula_complexity_score
            
        except Exception as e:
            self.logger.warning(f"Error creating file overview: {e}")
            overview['error'] = str(e)
        
        return overview
    
    def _create_executive_summary(self, context: AnalysisContext) -> str:
        """Create executive summary of analysis
        
        Args:
            context: AnalysisContext with module results
            
        Returns:
            String with executive summary
        """
        try:
            summary_parts = []
            
            # File basics
            file_name = context.file_path.name
            file_size = context.file_metadata.file_size_mb
            summary_parts.append(f"Analysis of {file_name} ({file_size:.1f}MB)")
            
            # Structure summary
            structure_result = context.get_module_result("structure_mapper")
            if structure_result and structure_result.data:
                worksheet_count = structure_result.data.worksheet_count
                summary_parts.append(f"Contains {worksheet_count} worksheets")
                
                if structure_result.data.chart_count > 0:
                    summary_parts.append(f"{structure_result.data.chart_count} charts")
                
                if structure_result.data.pivot_table_count > 0:
                    summary_parts.append(f"{structure_result.data.pivot_table_count} pivot tables")
            
            # Data quality summary
            data_result = context.get_module_result("data_profiler")
            if data_result and data_result.data:
                quality_score = data_result.data.data_quality_score
                if quality_score > 0.8:
                    quality_desc = "high"
                elif quality_score > 0.6:
                    quality_desc = "good"
                elif quality_score > 0.4:
                    quality_desc = "moderate"
                else:
                    quality_desc = "low"
                
                summary_parts.append(f"Data quality: {quality_desc} ({quality_score:.1%})")
            
            # Complexity summary
            formula_result = context.get_module_result("formula_analyzer")
            if formula_result and formula_result.data:
                formula_count = formula_result.data.total_formulas
                if formula_count > 1000:
                    complexity_desc = "high complexity"
                elif formula_count > 100:
                    complexity_desc = "moderate complexity"
                elif formula_count > 10:
                    complexity_desc = "low complexity"
                else:
                    complexity_desc = "minimal formulas"
                
                summary_parts.append(f"Formula analysis: {formula_count} formulas ({complexity_desc})")
            
            # Health issues
            health_result = context.get_module_result("health_checker")
            if health_result and health_result.data:
                if not health_result.data.file_accessible:
                    summary_parts.append("⚠️ File accessibility issues detected")
                elif health_result.data.security_issues:
                    summary_parts.append(f"⚠️ {len(health_result.data.security_issues)} security considerations")
            
            # Combine into summary
            if len(summary_parts) > 1:
                summary = ". ".join(summary_parts) + "."
            else:
                summary = f"Analysis completed for {file_name}."
            
            # Limit to reasonable length
            if len(summary) > 500:
                summary = summary[:497] + "..."
            
            return summary
            
        except Exception as e:
            self.logger.warning(f"Error creating executive summary: {e}")
            return f"Executive summary generation failed: {e}"
    
    def _collect_detailed_analysis(self, context: AnalysisContext) -> Dict[str, Any]:
        """Collect detailed analysis from all modules
        
        Args:
            context: AnalysisContext with module results
            
        Returns:
            Dict with detailed analysis data
        """
        detailed_analysis = {}
        
        # Collect results from each module
        module_names = [
            "health_checker", "structure_mapper", "data_profiler", 
            "formula_analyzer", "visual_cataloger", "connection_inspector", 
            "pivot_intelligence"
        ]
        
        for module_name in module_names:
            try:
                result = context.get_module_result(module_name)
                if result:
                    detailed_analysis[module_name] = {
                        'status': result.status.value,
                        'quality_score': result.quality_score,
                        'data_available': result.has_data,
                        'execution_time_seconds': result.execution_time,
                        'memory_usage_mb': result.memory_usage,
                        'error_count': len(result.errors),
                        'warning_count': len(result.warnings)
                    }
                    
                    # Include summary of actual data
                    if result.has_data and result.data:
                        detailed_analysis[module_name]['data_summary'] = self._summarize_module_data(
                            module_name, result.data
                        )
                else:
                    detailed_analysis[module_name] = {
                        'status': 'not_executed',
                        'data_available': False
                    }
            
            except Exception as e:
                self.logger.warning(f"Error collecting analysis for {module_name}: {e}")
                detailed_analysis[module_name] = {
                    'status': 'error',
                    'error': str(e)
                }
        
        return detailed_analysis
    
    def _summarize_module_data(self, module_name: str, data: Any) -> Dict[str, Any]:
        """Create summary of module-specific data
        
        Args:
            module_name: Name of the module
            data: Module data to summarize
            
        Returns:
            Dict with data summary
        """
        summary = {}
        
        try:
            if module_name == "health_checker":
                summary = {
                    'file_accessible': getattr(data, 'file_accessible', False),
                    'corruption_detected': getattr(data, 'corruption_detected', False),
                    'security_issue_count': len(getattr(data, 'security_issues', []))
                }
            
            elif module_name == "structure_mapper":
                summary = {
                    'worksheet_count': getattr(data, 'worksheet_count', 0),
                    'named_ranges_count': len(getattr(data, 'named_ranges', {})),
                    'chart_count': getattr(data, 'chart_count', 0),
                    'has_relationships': bool(getattr(data, 'sheet_relationships', {}))
                }
            
            elif module_name == "data_profiler":
                summary = {
                    'sheets_profiled': len(getattr(data, 'sheet_profiles', {})),
                    'data_quality_score': getattr(data, 'data_quality_score', 0.0),
                    'outliers_detected': getattr(data, 'outliers_detected', 0),
                    'patterns_found': len(getattr(data, 'patterns_found', []))
                }
            
            elif module_name == "formula_analyzer":
                summary = {
                    'total_formulas': getattr(data, 'total_formulas', 0),
                    'complexity_score': getattr(data, 'formula_complexity_score', 0.0),
                    'circular_references': len(getattr(data, 'circular_references', [])),
                    'external_references': len(getattr(data, 'external_references', []))
                }
            
            # Add summaries for other modules as needed
            else:
                # Generic summary for other modules
                if hasattr(data, '__dict__'):
                    summary = {key: str(value)[:100] for key, value in data.__dict__.items()}
                else:
                    summary = {'type': str(type(data)), 'content': str(data)[:100]}
        
        except Exception as e:
            summary = {'error': f"Failed to summarize {module_name} data: {e}"}
        
        return summary
    
    def _generate_recommendations(self, context: AnalysisContext) -> List[str]:
        """Generate actionable recommendations
        
        Args:
            context: AnalysisContext with module results
            
        Returns:
            List of recommendation strings
        """
        recommendations = []
        
        try:
            # Health-based recommendations
            health_result = context.get_module_result("health_checker")
            if health_result and health_result.data:
                if not health_result.data.file_accessible:
                    recommendations.append("File accessibility issues prevent full analysis - check file integrity")
                
                if health_result.data.macro_enabled:
                    recommendations.append("File contains macros - review for security before enabling")
                
                if health_result.data.file_size_mb > 50:
                    recommendations.append("Large file size detected - consider optimizing or splitting for better performance")
            
            # Data quality recommendations
            data_result = context.get_module_result("data_profiler")
            if data_result and data_result.data:
                if data_result.data.data_quality_score < 0.6:
                    recommendations.append("Data quality issues detected - review null values and data consistency")
                
                if data_result.data.duplicate_rows > 100:
                    recommendations.append("High number of duplicate rows detected - consider data deduplication")
                
                if data_result.data.outliers_detected > 50:
                    recommendations.append("Many outliers detected - review data for potential errors or validate extreme values")
            
            # Formula complexity recommendations
            formula_result = context.get_module_result("formula_analyzer")
            if formula_result and formula_result.data:
                if formula_result.data.total_formulas > 1000:
                    recommendations.append("High formula count detected - consider optimization for better performance")
                
                if formula_result.data.circular_references:
                    recommendations.append("Circular references detected - resolve to prevent calculation errors")
                
                if formula_result.data.external_references:
                    recommendations.append("External references detected - ensure linked files are available")
            
            # Structure recommendations
            structure_result = context.get_module_result("structure_mapper")
            if structure_result and structure_result.data:
                if structure_result.data.worksheet_count > 20:
                    recommendations.append("Many worksheets detected - consider organizing or consolidating")
                
                if len(structure_result.data.hidden_sheets) > 0:
                    recommendations.append("Hidden sheets detected - review for necessary content or security considerations")
            
            # Generic performance recommendations
            if context.file_metadata.file_size_mb > 100:
                recommendations.append("Consider enabling parallel processing for large file analysis")
            
            # Ensure we have at least one recommendation
            if not recommendations:
                recommendations.append("File analysis completed successfully - no immediate issues detected")
        
        except Exception as e:
            self.logger.warning(f"Error generating recommendations: {e}")
            recommendations.append(f"Recommendation generation failed: {e}")
        
        return recommendations
    
    def _create_ai_navigation_guide(self, context: AnalysisContext) -> Dict[str, Any]:
        """Create AI-friendly navigation guide
        
        Args:
            context: AnalysisContext with module results
            
        Returns:
            Dict with AI navigation information
        """
        navigation_guide = {
            'quick_facts': {},
            'data_locations': {},
            'analysis_confidence': {},
            'question_answers': []
        }
        
        try:
            # Quick facts for AI consumption
            structure_result = context.get_module_result("structure_mapper")
            if structure_result and structure_result.data:
                navigation_guide['quick_facts'] = {
                    'total_worksheets': structure_result.data.worksheet_count,
                    'worksheet_names': structure_result.data.worksheet_names,
                    'has_charts': structure_result.data.chart_count > 0,
                    'has_pivot_tables': structure_result.data.pivot_table_count > 0,
                    'has_hidden_sheets': len(structure_result.data.hidden_sheets) > 0
                }
            
            # Data locations for targeted analysis
            data_result = context.get_module_result("data_profiler")
            if data_result and data_result.data:
                data_locations = {}
                for sheet_name, profile in data_result.data.sheet_profiles.items():
                    if isinstance(profile, dict) and profile.get('status') == 'success':
                        data_locations[sheet_name] = {
                            'has_data': True,
                            'row_count': profile.get('row_count', 0),
                            'column_count': profile.get('column_count', 0),
                            'data_quality': profile.get('quality_score', 0.0)
                        }
                navigation_guide['data_locations'] = data_locations
            
            # Analysis confidence levels
            for module_name in ["health_checker", "structure_mapper", "data_profiler", "formula_analyzer"]:
                result = context.get_module_result(module_name)
                if result:
                    navigation_guide['analysis_confidence'][module_name] = {
                        'confidence': result.validation.confidence_level.value if result.validation else 'unknown',
                        'quality_score': result.quality_score,
                        'completed': result.is_successful
                    }
            
            # Common Q&A for AI assistance
            navigation_guide['question_answers'] = [
                {
                    'question': 'What worksheets contain actual data?',
                    'answer_location': 'data_locations',
                    'answer_key': 'sheets with has_data=true'
                },
                {
                    'question': 'Are there any data quality issues?',
                    'answer_location': 'detailed_analysis.data_profiler.data_summary',
                    'answer_key': 'data_quality_score'
                },
                {
                    'question': 'What charts and visualizations exist?',
                    'answer_location': 'detailed_analysis.visual_cataloger',
                    'answer_key': 'charts, images, shapes'
                },
                {
                    'question': 'Are there complex formulas or calculations?',
                    'answer_location': 'detailed_analysis.formula_analyzer',
                    'answer_key': 'total_formulas, complexity_score'
                }
            ]
        
        except Exception as e:
            self.logger.warning(f"Error creating AI navigation guide: {e}")
            navigation_guide['error'] = str(e)
        
        return navigation_guide
    
    def _create_metadata(self, context: AnalysisContext) -> Dict[str, Any]:
        """Create analysis metadata
        
        Args:
            context: AnalysisContext with analysis information
            
        Returns:
            Dict with metadata
        """
        try:
            analysis_summary = context.get_analysis_summary()
            
            metadata = {
                'analysis_version': '1.0',
                'timestamp': context.analysis_start_time,
                'file_path': str(context.file_path),
                'file_size_mb': context.file_metadata.file_size_mb,
                'processing_time_seconds': analysis_summary['timing']['elapsed_seconds'],
                'modules_executed': analysis_summary['progress']['module_names']['completed'],
                'modules_failed': analysis_summary['progress']['module_names']['failed'],
                'overall_success_rate': len(analysis_summary['progress']['module_names']['completed']) / 
                    (len(analysis_summary['progress']['module_names']['completed']) + 
                     len(analysis_summary['progress']['module_names']['failed'])) 
                    if (len(analysis_summary['progress']['module_names']['completed']) + 
                        len(analysis_summary['progress']['module_names']['failed'])) > 0 else 0.0,
                'resource_usage': analysis_summary['resources']
            }
            
            return metadata
        
        except Exception as e:
            self.logger.warning(f"Error creating metadata: {e}")
            return {'error': str(e)}
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Documentation synthesis is lightweight
        
        Args:
            context: AnalysisContext (unused)
            
        Returns:
            float: Low complexity multiplier
        """
        return 0.5  # Documentation synthesis is fast
    
    def _count_processed_items(self, data: Dict[str, Any]) -> int:
        """Count documentation sections created
        
        Args:
            data: Documentation data
            
        Returns:
            int: Number of sections created
        """
        return len([section for section in data.values() if section])


# Legacy compatibility
def create_doc_synthesizer(config: dict = None) -> DocSynthesizer:
    """Factory function for backward compatibility"""
    synthesizer = DocSynthesizer()
    if config:
        synthesizer.configure(config)
    return synthesizer
