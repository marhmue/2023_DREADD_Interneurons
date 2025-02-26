# 2023_ImARef

This Image Analysis Reformatting repository contains MATLAB scripts which extract outputs from the QUINT pipeline or GraphPad Prism statistics sheets, perform additional statistical calculations, and return re-formatted summary data for subsequent visualisation and evaluation.

The first MATLAB script, ImARef_Core.m, processes raw QUINT outputs. The program collates and re-formats technical and experimental replicate data for easy import to GraphPad Prism. Statistics sheets can then be generated using the Sidak correction for multiple comparisons, then loaded into the second MATLAB script (ImARef_HeatMaps.m). This second script returns re-configured overall and pairwise comparisons, which are formatted for export and heat map generation in GraphPad Prism, as well as statistical details and metadata. 

Custom blank GraphPad Prism templates (formatted for compatibility with MATLAB outputs) as well as custom .groovy and .ijm scripts (used as a part of the QUINT pipeline) can be provided upon reasonable request.
