"""
External connections mapping module
Framework-compatible placeholder with basic connection detection.
"""

from typing import List, Optional, Dict, Any

from ..core.base_analyzer import BaseAnalyzer
from ..core.analysis_context import AnalysisContext
from ..core.module_result import ConnectionData, ValidationResult, ConfidenceLevel
from ..utils.error_handler import ExcelAnalysisError, ErrorSeverity, ErrorCategory


class ConnectionInspector(BaseAnalyzer):
    """Framework-compatible connection inspector with basic implementation"""
    
    def __init__(self, name: str = "connection_inspector", dependencies: Optional[List[str]] = None):
        super().__init__(name, dependencies or ["health_checker"])
    
    def _perform_analysis(self, context: AnalysisContext) -> ConnectionData:
        """Perform basic external connection analysis
        
        Args:
            context: AnalysisContext with workbook access
            
        Returns:
            ConnectionData with connection inventory
        """
        try:
            # Connection inventories
            external_connections = []
            linked_workbooks = []
            database_connections = []
            web_queries = []
            refresh_settings = {}
            security_assessment = {}
            
            with context.get_workbook_access().get_workbook() as wb:
                # Check for external connections in workbook
                external_connections = self._detect_external_connections(wb)
                linked_workbooks = self._detect_linked_workbooks(wb)
                database_connections = self._detect_database_connections(wb)
                web_queries = self._detect_web_queries(wb)
                refresh_settings = self._analyze_refresh_settings(wb)
                security_assessment = self._assess_security_risks(
                    external_connections, database_connections, web_queries
                )
            
            return ConnectionData(
                external_connections=external_connections,
                linked_workbooks=linked_workbooks,
                database_connections=database_connections,
                web_queries=web_queries,
                refresh_settings=refresh_settings,
                security_assessment=security_assessment
            )
            
        except Exception as e:
            # Return minimal data rather than failing
            self.logger.error(f"Connection inspection failed: {e}")
            return ConnectionData(
                external_connections=[],
                linked_workbooks=[],
                database_connections=[],
                web_queries=[],
                refresh_settings={},
                security_assessment={'error': str(e)}
            )
    
    def _validate_result(self, data: ConnectionData, context: AnalysisContext) -> ValidationResult:
        """Validate connection inspection results
        
        Args:
            data: ConnectionData to validate
            context: AnalysisContext for validation
            
        Returns:
            ValidationResult with quality metrics
        """
        validation_notes = []
        
        # Completeness based on connection detection
        total_connections = (
            len(data.external_connections) + len(data.linked_workbooks) +
            len(data.database_connections) + len(data.web_queries)
        )
        
        # Most workbooks don't have many connections, so finding none is often correct
        if total_connections > 5:
            completeness = 0.9  # High connection activity
        elif total_connections > 0:
            completeness = 0.7  # Some connections found
        else:
            completeness = 0.6  # No connections (often valid)
        
        # Accuracy - assume good for basic detection
        accuracy = 0.8
        
        # Consistency checks
        consistency = 0.9
        if 'error' in data.security_assessment:
            consistency -= 0.3
            validation_notes.append("Security assessment failed")
        
        # Confidence based on connection complexity
        if total_connections > 10:
            confidence = ConfidenceLevel.MEDIUM
        elif total_connections > 2:
            confidence = ConfidenceLevel.LOW
        else:
            confidence = ConfidenceLevel.UNCERTAIN
        
        if len(data.database_connections) > 0:
            validation_notes.append("Database connections detected")
        if len(data.web_queries) > 0:
            validation_notes.append("Web queries detected")
        
        return ValidationResult(
            completeness_score=completeness,
            accuracy_score=accuracy,
            consistency_score=max(0.0, consistency),
            confidence_level=confidence,
            validation_notes=validation_notes
        )
    
    def _detect_external_connections(self, workbook) -> List[Dict[str, Any]]:
        """Detect external connections in workbook
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            List of external connection information
        """
        connections = []
        
        try:
            # Check workbook for external connections
            # Note: openpyxl has limited support for external connections
            # This is a basic placeholder implementation
            
            if hasattr(workbook, 'external_connections'):
                for i, conn in enumerate(workbook.external_connections):
                    conn_info = {
                        'id': f"ext_conn_{i}",
                        'type': 'external_connection',
                        'description': str(conn)[:200]  # Truncate long descriptions
                    }
                    connections.append(conn_info)
        
        except Exception as e:
            self.logger.warning(f"Error detecting external connections: {e}")
        
        return connections
    
    def _detect_linked_workbooks(self, workbook) -> List[str]:
        """Detect linked workbooks through formula references
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            List of linked workbook names
        """
        linked_workbooks = set()
        
        try:
            # Sample cells to find external workbook references
            max_cells_check = self.config.get("max_cell_checks", 1000)
            checked = 0
            
            for ws in workbook.worksheets:
                if checked >= max_cells_check:
                    break
                
                for row in ws.iter_rows():
                    if checked >= max_cells_check:
                        break
                    
                    for cell in row:
                        if checked >= max_cells_check:
                            break
                        
                        if (cell.value and isinstance(cell.value, str) and 
                            cell.value.startswith('=') and '[' in cell.value):
                            
                            # Extract workbook names from formulas like [workbook.xlsx]Sheet1!A1
                            import re
                            matches = re.findall(r'\[([^\]]+)\]', cell.value)
                            for match in matches:
                                if match.endswith(('.xlsx', '.xlsm', '.xls')):
                                    linked_workbooks.add(match)
                        
                        checked += 1
        
        except Exception as e:
            self.logger.warning(f"Error detecting linked workbooks: {e}")
        
        return list(linked_workbooks)
    
    def _detect_database_connections(self, workbook) -> List[Dict[str, Any]]:
        """Detect database connections
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            List of database connection information
        """
        db_connections = []
        
        try:
            # Placeholder for database connection detection
            # In a full implementation, this would check for:
            # - ODBC connections
            # - OLE DB connections
            # - SQL Server connections
            # - etc.
            
            # For now, just create placeholder structure
            if hasattr(workbook, 'connections'):
                for i, conn in enumerate(workbook.connections):
                    if 'database' in str(conn).lower() or 'sql' in str(conn).lower():
                        db_info = {
                            'id': f"db_conn_{i}",
                            'type': 'database',
                            'connection_string': str(conn)[:100],  # Truncate for security
                            'driver': 'unknown'
                        }
                        db_connections.append(db_info)
        
        except Exception as e:
            self.logger.warning(f"Error detecting database connections: {e}")
        
        return db_connections
    
    def _detect_web_queries(self, workbook) -> List[Dict[str, Any]]:
        """Detect web queries and web connections
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            List of web query information
        """
        web_queries = []
        
        try:
            # Placeholder for web query detection
            # In a full implementation, this would check for:
            # - Web query files (.iqy)
            # - Power Query connections
            # - Web service connections
            # - XML data imports
            
            # Basic detection by looking for web-like connection strings
            if hasattr(workbook, 'connections'):
                for i, conn in enumerate(workbook.connections):
                    conn_str = str(conn).lower()
                    if any(keyword in conn_str for keyword in ['http', 'web', 'url', 'xml']):
                        web_info = {
                            'id': f"web_query_{i}",
                            'type': 'web_query',
                            'url': conn_str[:200],  # Truncate long URLs
                            'method': 'unknown'
                        }
                        web_queries.append(web_info)
        
        except Exception as e:
            self.logger.warning(f"Error detecting web queries: {e}")
        
        return web_queries
    
    def _analyze_refresh_settings(self, workbook) -> Dict[str, Any]:
        """Analyze refresh settings for connections
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            Dict with refresh setting information
        """
        refresh_settings = {
            'refresh_on_open': False,
            'background_refresh': False,
            'refresh_interval_minutes': None,
            'last_refresh': None
        }
        
        try:
            # Placeholder for refresh settings analysis
            # This would examine workbook properties and connection settings
            # to determine refresh behavior
            
            # Basic implementation - check workbook properties
            if hasattr(workbook, 'properties'):
                props = workbook.properties
                refresh_settings['last_refresh'] = str(props.modified) if props.modified else None
        
        except Exception as e:
            self.logger.warning(f"Error analyzing refresh settings: {e}")
        
        return refresh_settings
    
    def _assess_security_risks(self, external_connections: List[Dict], 
                             database_connections: List[Dict], 
                             web_queries: List[Dict]) -> Dict[str, Any]:
        """Assess security risks from connections
        
        Args:
            external_connections: List of external connections
            database_connections: List of database connections
            web_queries: List of web queries
            
        Returns:
            Dict with security assessment
        """
        assessment = {
            'risk_level': 'low',
            'risk_factors': [],
            'recommendations': []
        }
        
        try:
            risk_score = 0
            
            # Assess risk factors
            if len(external_connections) > 5:
                assessment['risk_factors'].append('High number of external connections')
                risk_score += 2
            
            if len(database_connections) > 0:
                assessment['risk_factors'].append('Database connections present')
                assessment['recommendations'].append('Review database connection security')
                risk_score += 3
            
            if len(web_queries) > 0:
                assessment['risk_factors'].append('Web queries present')
                assessment['recommendations'].append('Verify web query URLs are trusted')
                risk_score += 2
            
            # Check for potentially risky connection patterns
            all_connections = external_connections + database_connections + web_queries
            for conn in all_connections:
                conn_str = str(conn).lower()
                if any(keyword in conn_str for keyword in ['password', 'pwd', 'secret']):
                    assessment['risk_factors'].append('Credentials may be stored in connections')
                    assessment['recommendations'].append('Use Windows Authentication where possible')
                    risk_score += 4
            
            # Determine overall risk level
            if risk_score >= 6:
                assessment['risk_level'] = 'high'
            elif risk_score >= 3:
                assessment['risk_level'] = 'medium'
            else:
                assessment['risk_level'] = 'low'
            
            assessment['risk_score'] = risk_score
        
        except Exception as e:
            self.logger.warning(f"Error assessing security risks: {e}")
            assessment['error'] = str(e)
        
        return assessment
    
    def estimate_complexity(self, context: AnalysisContext) -> float:
        """Estimate complexity for connection analysis
        
        Args:
            context: AnalysisContext with file metadata
            
        Returns:
            float: Complexity multiplier
        """
        base_complexity = super().estimate_complexity(context)
        
        # Connection analysis is relatively light
        return base_complexity * 0.8
    
    def _count_processed_items(self, data: ConnectionData) -> int:
        """Count connections processed
        
        Args:
            data: ConnectionData result
            
        Returns:
            int: Number of connections processed
        """
        return (
            len(data.external_connections) + len(data.linked_workbooks) +
            len(data.database_connections) + len(data.web_queries)
        )


# Legacy compatibility
def create_connection_inspector(config: dict = None) -> ConnectionInspector:
    """Factory function for backward compatibility"""
    inspector = ConnectionInspector()
    if config:
        inspector.configure(config)
    return inspector
