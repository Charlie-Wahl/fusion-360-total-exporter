# Author - Justin Nesselrotte
# Modified - IGES and DXF exports removed
# Description - Export all Fusion designs and projects (STEP, STL, Fusion Archive only)

from __future__ import with_statement

import adsk.core, adsk.fusion, adsk.cam, traceback
from logging import Logger, FileHandler, Formatter
import os
import re


class TotalExport(object):
    def __init__(self, app):
        self.app = app
        self.ui = self.app.userInterface
        self.data = self.app.data
        self.documents = self.app.documents
        self.log = Logger("Fusion 360 Total Export")
        self.num_issues = 0
        self.was_cancelled = False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def run(self, context):
        self.ui.messageBox(
            "Searching for and exporting files will take a while.\n\n"
            "Fusion must open and close every file in the main thread.\n\n"
            "Go grab a coffee ☕"
        )

        output_path = self._ask_for_output_path()
        if output_path is None:
            return

        file_handler = FileHandler(os.path.join(output_path, 'output.log'))
        file_handler.setFormatter(Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.log.addHandler(file_handler)

        self.log.info("Starting export!")
        self._export_data(output_path)
        self.log.info("Done exporting!")

        if self.was_cancelled:
            self.ui.messageBox("Cancelled!")
        elif self.num_issues > 0:
            self.ui.messageBox(
                f"The exporting process ran into {self.num_issues} issue"
                f"{'s' if self.num_issues > 1 else ''}. Check the log."
            )
        else:
            self.ui.messageBox("Export finished successfully!")

    def _export_data(self, output_path):
        progress_dialog = self.ui.createProgressDialog()
        progress_dialog.show("Exporting data!", "", 0, 1, 1)

        all_hubs = self.data.dataHubs
        for hub_index in range(all_hubs.count):
            hub = all_hubs.item(hub_index)
            self.log.info(f'Exporting hub "{hub.name}"')

            all_projects = hub.dataProjects
            for project_index in range(all_projects.count):
                project = all_projects.item(project_index)
                self.log.info(f'Exporting project "{project.name}"')

                files = self._get_files_for(project.rootFolder)

                progress_dialog.message = (
                    f"Hub: {hub_index + 1} of {all_hubs.count}\n"
                    f"Project: {project_index + 1} of {all_projects.count}\n"
                    "Exporting design %v of %m"
                )
                progress_dialog.maximumValue = len(files)
                progress_dialog.reset()

                for i, file in enumerate(files):
                    if progress_dialog.wasCancelled:
                        self.was_cancelled = True
                        return

                    progress_dialog.progressValue = i + 1
                    self._write_data_file(output_path, file)

    def _ask_for_output_path(self):
        dialog = self.ui.createFolderDialog()
        dialog.title = "Select export folder"
        if dialog.showDialog() != adsk.core.DialogResults.DialogOK:
            return None
        return dialog.folder

    def _get_files_for(self, folder):
        files = list(folder.dataFiles)
        for sub in folder.dataFolders:
            files.extend(self._get_files_for(sub))
        return files

    def _write_data_file(self, root_folder, file: adsk.core.DataFile):
        if file.fileExtension not in ("f3d", "f3z"):
            return

        try:
            document = self.documents.open(file)
            document.activate()
        except:
            self.num_issues += 1
            self.log.exception(f"Failed opening {file.name}")
            return

        try:
            file_folder = file.parentFolder
            path_parts = [self._name(file_folder.name)]

            while file_folder.parentFolder:
                file_folder = file_folder.parentFolder
                path_parts.insert(0, self._name(file_folder.name))

            project = file_folder.parentProject
            hub = project.parentHub

            base_path = self._take(
                root_folder,
                f"Hub {self._name(hub.name)}",
                f"Project {self._name(project.name)}",
                *path_parts
            )

            fusion_doc = adsk.fusion.FusionDocument.cast(document)
            design = fusion_doc.design
            export_mgr = design.exportManager

            archive_path = os.path.join(base_path, self._name(file.name))
            options = export_mgr.createFusionArchiveExportOptions(archive_path)
            export_mgr.execute(options)

            self._write_component(base_path, design.rootComponent)

        except:
            self.num_issues += 1
            self.log.exception(f"Failed processing {file.name}")
        finally:
            document.close(False)

    def _write_component(self, base_path, component: adsk.fusion.Component):
        output_path = os.path.join(base_path, self._name(component.name))

        self._write_step(output_path, component)
        self._write_stl(output_path, component)

        for occ in component.occurrences:
            self._write_component(output_path, occ.component)

    def _write_step(self, output_path, component):
        file_path = output_path + ".stp"
        if os.path.exists(file_path):
            return

        options = component.parentDesign.exportManager.createSTEPExportOptions(
            output_path, component
        )
        component.parentDesign.exportManager.execute(options)

    def _write_stl(self, output_path, component):
        file_path = output_path + ".stl"
        export_mgr = component.parentDesign.exportManager

        try:
            options = export_mgr.createSTLExportOptions(component, output_path)
            export_mgr.execute(options)
        except:
            pass

        for body in component.bRepBodies:
            self._write_stl_body(os.path.join(output_path, body.name), body)

        for body in component.meshBodies:
            self._write_stl_body(os.path.join(output_path, body.name), body)

    def _write_stl_body(self, output_path, body):
        file_path = output_path + ".stl"
        if os.path.exists(file_path):
            return

        try:
            options = body.parentComponent.parentDesign.exportManager.createSTLExportOptions(
                body, file_path
            )
            body.parentComponent.parentDesign.exportManager.execute(options)
        except:
            pass

    def _take(self, *path):
        out = os.path.join(*path)
        os.makedirs(out, exist_ok=True)
        return out

    def _name(self, name):
        name = re.sub(r'[^a-zA-Z0-9 \n\.]', '', name).strip()
        if name.lower().endswith(('.stp', '.stl')):
            name = name[:-4] + "_" + name[-3:]
        return name


def run(context):
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        with TotalExport(app) as exporter:
            exporter.run(context)
    except:
        ui.messageBox(f"Failed:\n{traceback.format_exc()}")
