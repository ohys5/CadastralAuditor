import math
from qgis.core import QgsGeometry, QgsPointXY, QgsWkbTypes, QgsSpatialIndex, Qgis

class GeometricMatcher:
    """
    Matches Cadastral feature with Current feature using Hausdorff Distance.
    """
    def __init__(self, cadastral_feat, current_feat):
        self.cadastral_geom = cadastral_feat.geometry()
        self.current_geom = current_feat.geometry()

    def get_hausdorff_distance(self):
        # Returns the maximum distance between the two geometries
        return self.current_geom.hausdorffDistance(self.cadastral_geom)

class BestFitSolver:
    """
    Calculates the optimal shift to align centroids and checks residual error.
    """
    def __init__(self, cadastral_feat, current_feat):
        self.cadastral_geom = cadastral_feat.geometry()
        self.current_geom = current_feat.geometry()
    
    def solve(self):
        # 1. Calculate Centroids
        if self.cadastral_geom.isEmpty() or self.current_geom.isEmpty():
            return 0.0, 0.0, 9999.0

        c_cad = self.cadastral_geom.centroid().asPoint()
        c_cur = self.current_geom.centroid().asPoint()
        
        # 2. Calculate Shift Vector (dx, dy)
        dx = c_cad.x() - c_cur.x()
        dy = c_cad.y() - c_cur.y()
        
        # 3. Create shifted geometry for testing
        shifted_geom = QgsGeometry(self.current_geom)
        shifted_geom.translate(dx, dy)
        
        # 4. Calculate Residual Error (Hausdorff distance after shift)
        residual = shifted_geom.hausdorffDistance(self.cadastral_geom)
        
        return dx, dy, residual

def SegmentAuditor(cadastral_feat, current_feat, tolerance=0.1):
    """
    Analyzes each segment of the current feature against the cadastral feature.
    Returns a summary string of the analysis.
    """
    current_geom = current_feat.geometry()
    cadastral_geom = cadastral_feat.geometry()
    
    # Convert to vertices to simulate segment analysis
    vertices = current_geom.vertices()
    
    issues = []
    for v in vertices:
        point_geom = QgsGeometry.fromPointXY(QgsPointXY(v))
        dist = point_geom.distance(cadastral_geom)
        
        if dist > tolerance:
            # Check if it is inside (gap) or outside (protrusion)
            if cadastral_geom.contains(point_geom):
                issues.append("이격")
            else:
                issues.append("돌출")
                
    return ", ".join(set(issues)) if issues else "부합"

class SmartGeometryComparator:
    def __init__(self, cadastral_feat, current_features):
        self.cadastral_feat = cadastral_feat
        self.current_features = current_features
        self.merged_geom = None
        self.mode = "Line Mode"

    def process(self):
        # 1. Combine and Merge Lines
        geoms = []
        for f in self.current_features:
            g = f.geometry()
            if g and not g.isEmpty():
                geoms.append(g)
        
        if not geoms:
            self.merged_geom = QgsGeometry()
            print("DEBUG: [Analyzer] No valid geometries in current_features")
        else:
            # Initialize with the first geometry
            combined = QgsGeometry(geoms[0])
            for g in geoms[1:]:
                combined = combined.combine(g)
            
            self.merged_geom = combined.mergeLines()
            
            # Fallback
            if not self.merged_geom or self.merged_geom.isEmpty():
                print("DEBUG: [Analyzer] mergeLines() returned empty, using combined geometry")
                self.merged_geom = combined
        
        print(f"DEBUG: [Analyzer] Merged Geom: {self.merged_geom.type()} (Empty: {self.merged_geom.isEmpty()})")
        
        # 2. Auto-Close Detection (5cm)
        if self.merged_geom.type() == QgsWkbTypes.LineGeometry:
            vertices = [v for v in self.merged_geom.vertices()]
            if vertices:
                start = vertices[0]
                end = vertices[-1]
                if start.distance(end) <= 0.05:
                    # Convert to Polygon
                    points = [QgsPointXY(v.x(), v.y()) for v in vertices]
                    if start.distance(end) > 0:
                        points.append(QgsPointXY(start.x(), start.y()))
                    
                    poly_geom = QgsGeometry.fromPolygonXY([points])
                    if poly_geom.isGeosValid():
                        self.merged_geom = poly_geom
                        self.mode = "Polygon Mode"

        result = {"mode": self.mode}
        cad_geom = self.cadastral_feat.geometry()

        if self.mode == "Polygon Mode":
            area_cur = self.merged_geom.area()
            area_cad = cad_geom.area()
            
            intersection = self.merged_geom.intersection(cad_geom)
            overlap_ratio = 0.0
            if area_cad > 0:
                overlap_ratio = (intersection.area() / area_cad) * 100.0
            
            result.update({
                "area_current": area_cur,
                "area_cadastral": area_cad,
                "overlap_ratio": overlap_ratio
            })
        else:
            # Line Mode: Explode and Find Target Segment
            segments = []
            
            print(f"DEBUG: [Analyzer] Cadastral Geom Type: {cad_geom.type()}, WKB: {cad_geom.wkbType()}")
            # Fix: Handle both Polygon and Line geometries to prevent TypeError
            if cad_geom.type() == QgsWkbTypes.PolygonGeometry:
                polys = cad_geom.asMultiPolygon() if cad_geom.isMultipart() else [cad_geom.asPolygon()]
                for poly in polys:
                    for ring in poly:
                        for i in range(len(ring) - 1):
                            segments.append(QgsGeometry.fromPolylineXY([ring[i], ring[i+1]]))
            elif cad_geom.type() == QgsWkbTypes.LineGeometry:
                lines = cad_geom.asMultiPolyline() if cad_geom.isMultipart() else [cad_geom.asPolyline()]
                for line in lines:
                    for i in range(len(line) - 1):
                        segments.append(QgsGeometry.fromPolylineXY([line[i], line[i+1]]))
            
            print(f"DEBUG: [Analyzer] Extracted segments: {len(segments)}")
            
            # Find closest segment (Target Segment)
            target_segment = min(segments, key=lambda s: self.merged_geom.hausdorffDistance(s), default=None)
            
            # Calculate Perpendicular Distance (Centroid to Segment)
            centroid = self.merged_geom.centroid()
            error_line = None
            perp_dist = -1.0
            
            if target_segment and not centroid.isEmpty():
                print(f"DEBUG: [Analyzer] Target Segment: {target_segment.asWkt()}")
                print(f"DEBUG: [Analyzer] Centroid: {centroid.asWkt()}")
                perp_dist = target_segment.distance(centroid)
                closest_pt = target_segment.nearestPoint(centroid)
                if not closest_pt.isEmpty():
                    error_line = QgsGeometry.fromPolylineXY([centroid.asPoint(), closest_pt.asPoint()])
                    print(f"DEBUG: [Analyzer] Error line created: {error_line.asWkt()}")
                else:
                    print("DEBUG: [Analyzer] Closest point is empty")
            else:
                print("DEBUG: [Analyzer] Fallback to Hausdorff (No target segment or empty centroid)")
                perp_dist = self.merged_geom.hausdorffDistance(cad_geom)
            
            print(f"DEBUG: [Analyzer] Calculated perp_dist: {perp_dist}")
            
            result.update({
                "max_discrepancy": perp_dist,
                "error_line": error_line
            })
            
        return result

class ParcelBasedAuditor:
    """
    Analyzes survey lines by intersecting them with cadastral parcels (Polygon).
    Performs Reference Matching and calculates Average Perpendicular Distance.
    """
    def __init__(self, cadastral_features, survey_feature):
        self.cadastral_features = cadastral_features
        self.survey_feature = survey_feature

    def run(self):
        results = []
        survey_geom = self.survey_feature.geometry()
        
        for cad_feat in self.cadastral_features:
            cad_geom = cad_feat.geometry()
            
            if not cad_geom.intersects(survey_geom):
                continue
                
            intersection = survey_geom.intersection(cad_geom)
            if intersection.isEmpty():
                continue
                
            segments = intersection.asMultiPolyline() if intersection.isMultipart() else [intersection.asPolyline()]
            
            for seg_points in segments:
                if len(seg_points) < 2: continue
                seg_geom = QgsGeometry.fromPolylineXY(seg_points)
                if seg_geom.length() == 0: continue
                
                ref_info = self.find_reference_line(seg_geom, cad_geom)
                
                if ref_info:
                    avg_dist = self.calculate_average_distance(seg_geom, ref_info['geometry'])
                    
                    pnu = str(cad_feat.id())
                    for field in ['jibun', 'pnu', 'JIBUN', 'PNU']:
                        if cad_feat.fieldNameIndex(field) != -1:
                            pnu = str(cad_feat[field])
                            break
                    
                    results.append({
                        "pnu": pnu,
                        "direction": ref_info['direction'],
                        "avg_error": avg_dist,
                        "max_error": seg_geom.hausdorffDistance(ref_info['geometry'])
                    })
        return results

    def find_reference_line(self, survey_segment, parcel_geom):
        boundaries = []
        polys = parcel_geom.asMultiPolygon() if parcel_geom.isMultipart() else [parcel_geom.asPolygon()]
        
        for poly in polys:
            for ring in poly:
                for i in range(len(ring) - 1):
                    boundaries.append(QgsGeometry.fromPolylineXY([ring[i], ring[i+1]]))
        
        survey_angle = self._get_angle(survey_segment)
        parcel_center = parcel_geom.centroid().asPoint()
        
        best_ref = None
        min_dist = float('inf')
        best_dir = "Unknown"
        
        for boundary in boundaries:
            b_angle = self._get_angle(boundary)
            diff = abs(survey_angle - b_angle)
            if diff > 180: diff = 360 - diff
            
            if diff < 30 or abs(diff - 180) < 30:
                dist = boundary.distance(survey_segment)
                if dist < min_dist:
                    min_dist = dist
                    best_ref = boundary
                    
                    mid = boundary.interpolate(boundary.length()/2).asPoint()
                    dx = mid.x() - parcel_center.x()
                    dy = mid.y() - parcel_center.y()
                    if abs(dx) > abs(dy):
                        best_dir = "동측" if dx > 0 else "서측"
                    else:
                        best_dir = "북측" if dy > 0 else "남측"
                        
        return {'geometry': best_ref, 'direction': best_dir} if best_ref else None

    def _get_angle(self, line):
        if line.isEmpty(): return 0
        p1 = line.vertexAt(0)
        p2 = line.vertexAt(line.numPoints()-1)
        return math.degrees(math.atan2(p2.y() - p1.y(), p2.x() - p1.x()))

class ConfidenceStringMatcher:
    """
    Adaptive Matching Engine using Confidence Region String Matching.
    Handles topology-aware logic (Closed vs Open) and calculates F-measure based on confidence bands.
    """
    def __init__(self, sigma=0.1):
        self.sigma = sigma
        self.band_width = sigma * 1.96 # 95% 신뢰구간 폭

    def clip_to_overlap(self, g1, g2, buff_width):
        b2 = g2.buffer(buff_width, 5)
        c1 = g1.intersection(b2)
        b1 = g1.buffer(buff_width, 5)
        c2 = g2.intersection(b1)
        return c1, c2

    def process_pair(self, survey_feat, cadastral_feat, th_shape=0.25, th_pos=0.15, mode='distance'):
        """
        현형선과 지적선의 기하학적 유사성을 신뢰 영역(Confidence Region) 기반으로 분석합니다.
        :param mode: 'distance' (거리 기반) 또는 'original' (기존 면적/신뢰구간 기반)
        """
        survey_geom = survey_feat.geometry()
        cadastral_geom = cadastral_feat.geometry()
        
        # Calculate robust shift using PointToLineAuditor (Average Deviation)
        p2l = PointToLineAuditor(cadastral_feat, survey_feat, densify_distance=1.0)
        p2l_res = p2l.process()
        dx, dy = p2l_res['avg_dx'], p2l_res['avg_dy']
        
        # ---------------------------------------------------------
        # MODE A: 거리 기반 (Distance Based)
        # ---------------------------------------------------------
        if mode == 'distance':
            # 1. 절대 거리 (MAE - 평균 거리 사용)
            dist_absolute = p2l_res['mae']
            
            # 2. 형상 거리 (Centroid Shift 후 Hausdorff)
            if survey_geom.isEmpty() or cadastral_geom.isEmpty():
                return {
                    "topology": "Invalid",
                    "score": 999.0,
                    "status": "Error",
                    "nd_cost": 999.0
                }

            geom_curr_shifted = QgsGeometry(survey_geom)
            geom_curr_shifted.translate(dx, dy)
            dist_shape = geom_curr_shifted.hausdorffDistance(cadastral_geom)
            
            # 판정
            # th_pos: 위치 오차 허용치 (m) -> dist_absolute와 비교
            # th_shape: 형상 오차 허용치 (m) -> dist_shape와 비교
            
            if dist_absolute <= th_pos:
                status = "부합 (Pass)"
            elif dist_shape <= th_shape:
                status = "위치정정 필요"
            else:
                status = "불부합"
                
            return {
                "topology": "Distance",
                "score": dist_absolute,      # 절대오차 (m)
                "status": status,
                "nd_cost": dist_shape,       # 형상오차 (m)
                "shift_x": dx,
                "shift_y": dy
            }

        # ---------------------------------------------------------
        # MODE B: 기존 면적/점수 기반 (Original / ND Cost)
        # ---------------------------------------------------------
        else:
            # Ensure LineString for Band Analysis (Convert Polygon to Boundary)
            if survey_geom.type() == QgsWkbTypes.PolygonGeometry:
                survey_geom = QgsGeometry(survey_geom) # Clone to avoid modifying original
                survey_geom.convertToType(QgsWkbTypes.LineGeometry)
            if cadastral_geom.type() == QgsWkbTypes.PolygonGeometry:
                cadastral_geom = QgsGeometry(cadastral_geom)
                cadastral_geom.convertToType(QgsWkbTypes.LineGeometry)

            # 1. [보완] 교차 검사 (Topology Check)
            is_crossed = False
            # Check for crossing (X-shape) which implies topological mismatch
            if survey_geom.crosses(cadastral_geom): 
                is_crossed = True

            # 2. [보완] 범위 일치화 (Projection Clipping)
            # 서로 겹치는 구간만 잘라내어 순수 형상 비교 준비
            clip_width = max(th_pos * 2.0, 1.0)
            seg_surv, seg_cad = self.clip_to_overlap(survey_geom, cadastral_geom, clip_width)
            
            if seg_surv.isEmpty() or seg_cad.isEmpty():
                return {
                    "topology": "No Overlap",
                    "score": 0.0,
                    "status": "불부합 (No Overlap)",
                    "nd_cost": 999.0,
                    "position_error": 999.0,
                    "shift_x": dx,
                    "shift_y": dy
                }

            # 3. 버퍼 생성 및 면적 계산 (신뢰 띠)
            buff_s = seg_surv.buffer(self.band_width, 8, Qgis.EndCapStyle.Round, Qgis.JoinStyle.Round, 2.0)
            buff_c = seg_cad.buffer(self.band_width, 8, Qgis.EndCapStyle.Round, Qgis.JoinStyle.Round, 2.0)
            
            area_union = buff_s.combine(buff_c).area()
            area_int = buff_s.intersection(buff_c).area()
            
            len_sum = seg_surv.length() + seg_cad.length()
            nd_cost = (area_union - area_int) / len_sum if len_sum > 0 else 1.0

            # 4. 위치 오차 계산 (Position Error)
            # 실제 중첩 폭 = area_int / average_length
            avg_len = len_sum / 2.0
            overlap_width = area_int / avg_len if avg_len > 0 else 0.0
            
            max_width = self.band_width * 2
            position_error = max_width - overlap_width
            if position_error < 0: position_error = 0.0
            
            # 5. [보완] 회전(Rotation) 여부 판단
            # 클리핑된 현형선의 양 끝점에서 지적선까지의 거리 차이 비교
            verts = [v for v in seg_surv.vertices()]
            is_rotated = False
            if verts:
                p_start = QgsGeometry.fromPointXY(QgsPointXY(verts[0].x(), verts[0].y()))
                p_end = QgsGeometry.fromPointXY(QgsPointXY(verts[-1].x(), verts[-1].y()))
                dist_start = p_start.distance(seg_cad)
                dist_end = p_end.distance(seg_cad)
                is_rotated = abs(dist_start - dist_end) > th_pos

            # 6. 종합 판정 (Matrix Decision)
            score = overlap_width / max_width if max_width > 0 else 0.0
            if score > 1.0: score = 1.0

            if is_crossed:
                status = "불부합 (Crossed)"
            elif nd_cost <= th_shape:
                if position_error <= th_pos:
                    status = "부합 (Pass)"
                else:
                    if is_rotated:
                        status = "회전 보정 필요"
                    else:
                        status = "위치정정 필요"
            else:
                status = "형상 불일치"
            
            return {
                "topology": "Band Analysis",
                "score": score,
                "status": status,
                "nd_cost": nd_cost,
                "position_error": position_error,
                "shift_x": dx,
                "shift_y": dy
            }

    def calculate_average_distance(self, seg, ref):
        p_start = QgsGeometry.fromPointXY(QgsPointXY(seg.vertexAt(0).x(), seg.vertexAt(0).y()))
        p_end = QgsGeometry.fromPointXY(QgsPointXY(seg.vertexAt(seg.numPoints()-1).x(), seg.vertexAt(seg.numPoints()-1).y()))
        return (ref.distance(p_start) + ref.distance(p_end)) / 2.0

class PointToLineAuditor:
    """
    P2L (Point-to-Line) Algorithm:
    1. Vertex Matching: Corners (>=45 deg) match to Cadastral Vertices.
    2. Contextual Matching: Short segments (<2m) inherit target from neighbors.
    3. Densification: Check every 2m.
    """
    def __init__(self, cadastral_feat, current_feat, densify_distance=2.0):
        self.cadastral_geom = cadastral_feat.geometry()
        self.current_geom = current_feat.geometry()
        self.densify_dist = densify_distance

    def process(self):
        # 1. Prepare Cadastral Data (Segments & Vertices)
        cad_segments_data = [] # List of (geometry, angle)
        cad_vertices = []
        
        # Robust extraction of lines from Polygon/Line geometries
        cad_parts = []
        if self.cadastral_geom.type() == QgsWkbTypes.PolygonGeometry:
            if self.cadastral_geom.isMultipart():
                for poly in self.cadastral_geom.asMultiPolygon():
                    cad_parts.extend(poly)
            else:
                cad_parts.extend(self.cadastral_geom.asPolygon())
        elif self.cadastral_geom.type() == QgsWkbTypes.LineGeometry:
            if self.cadastral_geom.isMultipart():
                cad_parts = self.cadastral_geom.asMultiPolyline()
            else:
                cad_parts = [self.cadastral_geom.asPolyline()]
        
        # Step 1: Reference Densification (for Vertices)
        # Apply densifyByDistance using the same distance as survey line
        densified_cad = self.cadastral_geom.densifyByDistance(self.densify_dist)
        for v in densified_cad.vertices():
            cad_vertices.append(QgsPointXY(v.x(), v.y()))
            
        for part in cad_parts:
            for i in range(len(part)):
                if i < len(part) - 1:
                    p1, p2 = part[i], part[i+1]
                    geom = QgsGeometry.fromPolylineXY([p1, p2])
                    # Calculate Azimuth (Angle)
                    angle = math.degrees(math.atan2(p2.y() - p1.y(), p2.x() - p1.x()))
                    cad_segments_data.append((geom, angle))

        # 2. Process Current Geometry
        cur_parts = []
        if self.current_geom.type() == QgsWkbTypes.PolygonGeometry:
            if self.current_geom.isMultipart():
                for poly in self.current_geom.asMultiPolygon():
                    for ring in poly:
                        cur_parts.append(ring)
            else:
                for ring in self.current_geom.asPolygon():
                    cur_parts.append(ring)
        elif self.current_geom.type() == QgsWkbTypes.LineGeometry:
            if self.current_geom.isMultipart():
                cur_parts = self.current_geom.asMultiPolyline()
            else:
                cur_parts = [self.current_geom.asPolyline()]

        max_dev = 0.0
        sum_sq = 0.0
        sum_dist = 0.0
        count = 0
        vectors = []
        sum_dx = 0.0
        sum_dy = 0.0
        
        for part in cur_parts:
            if len(part) < 2: continue
            
            # A. Identify Corners
            corners = [False] * len(part)
            for i in range(1, len(part) - 1):
                p_prev, p_curr, p_next = part[i-1], part[i], part[i+1]
                a1 = math.atan2(p_curr.y() - p_prev.y(), p_curr.x() - p_prev.x())
                a2 = math.atan2(p_next.y() - p_curr.y(), p_next.x() - p_curr.x())
                diff = math.degrees(abs(a1 - a2))
                if diff > 180: diff = 360 - diff
                if diff >= 45:
                    corners[i] = True
            
            # C. Iterative Check (Densified)
            for i in range(len(part) - 1):
                p_start, p_end = part[i], part[i+1]
                
                # Dynamic Target Selection: Find ALL parallel cadastral segments
                # This handles cases where the survey segment spans multiple cadastral segments
                curr_angle = math.degrees(math.atan2(p_end.y() - p_start.y(), p_end.x() - p_start.x()))
                
                candidates = []
                for geom, ang in cad_segments_data:
                    diff = abs(curr_angle - ang)
                    if diff > 180: diff = 360 - diff
                    if diff <= 30 or abs(diff - 180) <= 30: # Parallel check
                        candidates.append(geom)
                
                if candidates:
                    # Combine all parallel segments into one target geometry
                    target_geom = QgsGeometry.fromMultiPolylineXY([c.asPolyline() for c in candidates])
                else:
                    # Fallback to full cadastral geometry if no parallel found
                    target_geom = self.cadastral_geom
                
                seg_geom = QgsGeometry.fromPolylineXY([p_start, p_end])
                densified = seg_geom.densifyByDistance(self.densify_dist).asPolyline()
                
                # Process points (avoid double counting p_end unless last segment)
                points_to_check = densified if i == len(part) - 2 else densified[:-1]
                
                for j, pt in enumerate(points_to_check):
                    pt_geom = QgsGeometry.fromPointXY(pt)
                    is_corner = (j == 0 and corners[i])
                    if i == len(part) - 2 and j == len(points_to_check) - 1 and corners[i+1]:
                        is_corner = True
                        
                    if is_corner:
                        # Vertex Matching
                        near_pt = min(cad_vertices, key=lambda v: pt.sqrDist(v))
                        dist = math.sqrt(pt.sqrDist(near_pt))
                    else:
                        # Contextual Matching
                        # Step 3: Strict Perpendicular Projection
                        nearest_res = target_geom.nearestPoint(pt_geom)
                        near_pt = nearest_res.asPoint()
                        dist = pt_geom.distance(nearest_res)
            
                    if dist > max_dev: max_dev = dist
                    sum_sq += dist * dist
                    sum_dist += dist
                    count += 1
                    sum_dx += (near_pt.x() - pt.x())
                    sum_dy += (near_pt.y() - pt.y())
                    
                    if dist >= 0.1:
                        vectors.append(QgsGeometry.fromPolylineXY([pt, near_pt]))
                
        rmse = math.sqrt(sum_sq / count) if count > 0 else 0.0
        mae = sum_dist / count if count > 0 else 0.0
        avg_dx = sum_dx / count if count > 0 else 0.0
        avg_dy = sum_dy / count if count > 0 else 0.0
        
        multi_line = QgsGeometry.fromMultiPolylineXY([v.asPolyline() for v in vectors]) if vectors else QgsGeometry()
        
        return {
            "max_deviation": max_dev,
            "rmse": rmse,
            "mae": mae,
            "avg_dx": avg_dx,
            "avg_dy": avg_dy,
            "error_vectors": multi_line
        }

class TopologyAuditor:
    """
    Advanced Topology Normalization:
    1. Split Process (1:N): Split survey lines by cadastral boundaries.
    2. Group Process (N:1): Group segments referencing the same cadastral edge.
    3. Weighted Average Error: Calculate error for the group.
    """
    def __init__(self, cad_layer, cur_layer, transform=None):
        self.cad_layer = cad_layer
        self.cur_layer = cur_layer
        self.xform = transform

    def process(self):
        results = []
        # 1. Build Spatial Index for Cadastral Layer
        cad_index = QgsSpatialIndex(self.cad_layer.getFeatures())
        
        # Intermediate storage for grouping
        # Key: target_edge_id (str), Value: dict of stats
        groups = {}
        
        # 2. Iterate & Split
        for cur_feat in self.cur_layer.getFeatures():
            geom = cur_feat.geometry()
            if not geom: continue
            if self.xform:
                geom.transform(self.xform)
                
            bbox = geom.boundingBox()
            cids = cad_index.intersects(bbox)
            
            for cid in cids:
                cad_feat = self.cad_layer.getFeature(cid)
                cad_geom = cad_feat.geometry()
                
                if not cad_geom.intersects(geom): continue
                
                intersection = geom.intersection(cad_geom)
                if intersection.isEmpty(): continue
                
                parts = intersection.asMultiPolyline() if intersection.isMultipart() else [intersection.asPolyline()]
                
                for part in parts:
                    if len(part) < 2: continue
                    seg_geom = QgsGeometry.fromPolylineXY(part)
                    
                    # Find Target Edge
                    target = self.find_target_edge(seg_geom, cad_feat)
                    if target:
                        tid = target['id']
                        if tid not in groups:
                            groups[tid] = {
                                'pnu': self.get_pnu(cad_feat),
                                'w_error_sum': 0.0,
                                'len_sum': 0.0,
                                'vectors': [],
                                'fids': []
                            }
                        
                        # Calculate average distance for this segment
                        d1 = target['geom'].distance(QgsGeometry.fromPointXY(QgsPointXY(part[0])))
                        d2 = target['geom'].distance(QgsGeometry.fromPointXY(QgsPointXY(part[-1])))
                        avg_dist = (d1 + d2) / 2.0
                        
                        groups[tid]['w_error_sum'] += avg_dist * seg_geom.length()
                        groups[tid]['len_sum'] += seg_geom.length()
                        groups[tid]['fids'].append(cur_feat.id())
                        
                        # Create vector for visualization (midpoint to line)
                        mid = seg_geom.interpolate(seg_geom.length()/2)
                        proj = target['geom'].nearestPoint(mid)
                        vec = QgsGeometry.fromPolylineXY([mid.asPoint(), proj.asPoint()])
                        groups[tid]['vectors'].append(vec)

        # 3. Finalize Groups
        for tid, g in groups.items():
            if g['len_sum'] == 0: continue
            avg_error = g['w_error_sum'] / g['len_sum']
            
            combined_vector = QgsGeometry()
            if g['vectors']:
                combined_vector = QgsGeometry.fromMultiPolylineXY([v.asPolyline() for v in g['vectors']])
                
            results.append({
                'pnu': g['pnu'],
                'error': avg_error,
                'length': g['len_sum'],
                'vector': combined_vector,
                'fid': g['fids'][0] if g['fids'] else None
            })
            
        return results

    def get_pnu(self, feat):
        for field in ['jibun', 'pnu', 'JIBUN', 'PNU']:
            if feat.fieldNameIndex(field) != -1:
                return str(feat[field])
        return str(feat.id())

    def find_target_edge(self, seg_geom, cad_feat):
        edges = []
        geom = cad_feat.geometry()
        if geom.isMultipart():
            polys = geom.asMultiPolygon()
        else:
            polys = [geom.asPolygon()]
            
        idx = 0
        for poly in polys:
            for ring in poly:
                for i in range(len(ring)-1):
                    p1, p2 = ring[i], ring[i+1]
                    edge = QgsGeometry.fromPolylineXY([p1, p2])
                    edges.append((f"{cad_feat.id()}_{idx}", edge))
                    idx += 1
        
        seg_angle = self._get_angle(seg_geom)
        best = None
        min_dist = float('inf')
        
        for eid, edge in edges:
            edge_angle = self._get_angle(edge)
            diff = abs(seg_angle - edge_angle)
            if diff > 180: diff = 360 - diff
            
            if diff <= 30 or abs(diff - 180) <= 30:
                dist = edge.distance(seg_geom)
                if dist < min_dist:
                    min_dist = dist
                    best = {'id': eid, 'geom': edge}
        
        return best

    def _get_angle(self, line):
        if line.isEmpty(): return 0
        p1 = line.vertexAt(0)
        p2 = line.vertexAt(line.numPoints()-1)
        return math.degrees(math.atan2(p2.y() - p1.y(), p2.x() - p1.x()))