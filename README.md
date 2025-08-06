import zipfile
import os
import shutil

def create_icon_catalog_from_vssx(vssx_path: str, output_vsdx_path: str):
    tmp_dir = "tmp_vssx_extract"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    # Step 1: Extract the .vssx file
    with zipfile.ZipFile(vssx_path, 'r') as z:
        z.extractall(tmp_dir)

    masters_dir = os.path.join(tmp_dir, "visio", "masters")
    master_files = [f for f in os.listdir(masters_dir) if f.endswith(".xml")]

    # Step 2: Build a new vsdx zip
    with zipfile.ZipFile(output_vsdx_path, 'w', zipfile.ZIP_DEFLATED) as vsdx:

        # Write base content files
        vsdx.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.document.main+xml"/>
  <Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/>
</Types>""")

        vsdx.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document"
    Target="/visio/document.xml"/>
</Relationships>""")

        # Write document.xml with master references
        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
<VisioDocument xmlns="http://schemas.microsoft.com/visio/2010/visioDocument">
  <Pages>
    <Page ID="1" NameU="Page-1" Name="Page-1"/>
  </Pages>
  <Masters>
"""
        for i, master in enumerate(master_files):
            document_xml += f'    <Master ID="{i+1}" NameU="{master[:-4]}" Name="{master[:-4]}" MasterShortcut="{master}"/>\n'
        document_xml += "  </Masters>\n</VisioDocument>"
        vsdx.writestr("visio/document.xml", document_xml)

        # Step 3: Create a page and layout shapes (grid)
        cols = 5
        x_spacing = 2
        y_spacing = 2
        start_x, start_y = 2, 10

        page_contents = """<?xml version="1.0" encoding="UTF-8"?>
<PageContents xmlns="http://schemas.microsoft.com/visio/2010/page">
"""
        for idx, master in enumerate(master_files):
            row = idx // cols
            col = idx % cols
            pinx = start_x + col * x_spacing
            piny = start_y - row * y_spacing
            page_contents += f"""
  <Shape ID="{idx+1}" Master="{idx+1}" Type="Shape" Text="{master[:-4]}">
    <XForm>
      <PinX>{pinx}</PinX>
      <PinY>{piny}</PinY>
      <Width>1.5</Width>
      <Height>1.5</Height>
    </XForm>
  </Shape>"""
        page_contents += "\n</PageContents>"
        vsdx.writestr("visio/pages/page1.xml", page_contents)
        vsdx.writestr("visio/pages/_rels/page1.xml.rels", """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""")

        # Step 4: Inject all master shape XMLs
        for master in master_files:
            with open(os.path.join(masters_dir, master), "rb") as f:
                vsdx.writestr(f"visio/masters/{master}", f.read())

        rels_path = os.path.join(masters_dir, "_rels")
        if os.path.exists(rels_path):
            for rel_file in os.listdir(rels_path):
                with open(os.path.join(rels_path, rel_file), "rb") as f:
                    vsdx.writestr(f"visio/masters/_rels/{rel_file}", f.read())

    print(f"✅ Generated: {output_vsdx_path}")


# Example usage:
# Replace with your actual stencil path
create_icon_catalog_from_vssx("Azure_All.vssx", "blank.vsdx")
