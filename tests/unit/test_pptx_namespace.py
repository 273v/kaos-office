"""Unit tests for PresentationML namespace constants."""

from __future__ import annotations

from kaos_office.ooxml.namespace import (
    A_BU_AUTO_NUM,
    A_BU_CHAR,
    A_BU_NONE,
    A_P,
    A_R,
    A_T,
    A_TBL,
    C_CHART,
    C_SER,
    CHART_TYPE_TAGS,
    DGM,
    DGM_PT,
    DGM_REL_IDS,
    GD_CHART,
    GD_DIAGRAM,
    GD_TABLE,
    P_SP,
    P_SP_TREE,
    RT_CHART,
    RT_DIAGRAM_DATA,
    RT_SLIDE,
    A,
    C,
    P,
)


class TestPresentationMLConstants:
    """Test PresentationML namespace URI and tag constants."""

    def test_p_namespace(self):
        assert P == "http://schemas.openxmlformats.org/presentationml/2006/main"

    def test_c_namespace(self):
        assert C == "http://schemas.openxmlformats.org/drawingml/2006/chart"

    def test_dgm_namespace(self):
        assert DGM == "http://schemas.openxmlformats.org/drawingml/2006/diagram"

    def test_p_sp_tag(self):
        assert f"{{{P}}}sp" == P_SP

    def test_p_sp_tree_tag(self):
        assert f"{{{P}}}spTree" == P_SP_TREE

    def test_a_p_tag(self):
        assert f"{{{A}}}p" == A_P

    def test_a_r_tag(self):
        assert f"{{{A}}}r" == A_R

    def test_a_t_tag(self):
        assert f"{{{A}}}t" == A_T

    def test_a_bullet_tags(self):
        assert f"{{{A}}}buChar" == A_BU_CHAR
        assert f"{{{A}}}buAutoNum" == A_BU_AUTO_NUM
        assert f"{{{A}}}buNone" == A_BU_NONE

    def test_a_tbl_tag(self):
        assert f"{{{A}}}tbl" == A_TBL


class TestChartConstants:
    """Test chart namespace constants."""

    def test_chart_tag(self):
        assert f"{{{C}}}chart" == C_CHART

    def test_chart_ser_tag(self):
        assert f"{{{C}}}ser" == C_SER

    def test_chart_type_tags_not_empty(self):
        assert len(CHART_TYPE_TAGS) >= 10


class TestDiagramConstants:
    """Test diagram/SmartArt constants."""

    def test_dgm_pt_tag(self):
        assert f"{{{DGM}}}pt" == DGM_PT

    def test_dgm_rel_ids_tag(self):
        assert f"{{{DGM}}}relIds" == DGM_REL_IDS


class TestRelationshipTypes:
    """Test PresentationML relationship type URIs."""

    def test_rt_slide(self):
        assert "slide" in RT_SLIDE.lower()

    def test_rt_chart(self):
        assert "chart" in RT_CHART.lower()

    def test_rt_diagram_data(self):
        assert "diagramData" in RT_DIAGRAM_DATA


class TestGraphicDataURIs:
    """Test graphicData URI constants for shape identification."""

    def test_gd_table(self):
        assert "table" in GD_TABLE

    def test_gd_chart(self):
        assert "chart" in GD_CHART

    def test_gd_diagram(self):
        assert "diagram" in GD_DIAGRAM
