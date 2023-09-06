# 2023_DREADD_Interneurons

This repository contains MATLAB scripts which extract outputs from the QUINT pipeline or GraphPad Prism statistics sheets, perform additional statistical calculations, and return re-formatted summary data for subsequent import to GraphPad Prism.

The first MATLAB script quint_postprocessing_subdivisions.m processes raw QUINT outputs. The program collates technical replicate data (coronal sections from the same animal) in outputted rows 2-5, then furthermore collates experimental replicates (animals of the same genotype and DREADD condition) to re-format data per experimental condition and brain region for easy import to GraphPad Prism.

Once data are copied to GraphPad Prism, statistics can be run and exported in a new folder for subsequent import to the second and third MATLAB script.

The second MATLAB script, prism_anova_reformatting_sidak.m, re-formats outputted GraphPad Prism statistics sheets obtained using the Sidak correction for multiple comparisons. Resulting rows and colums are tailored for import to GraphPad Prism for heat map generation (constituting summary row and column values in the final figure), and to return detailed statistical information in columns for data reporting in publications.

The third MATLAB script, prism_anova_reformatting_pairwise_stats, re-formats independently outputted GraphPad Prism statistics sheets obtained using the two-stage linear setup of Benjamini, Krieger, and Yekutieli. It extracts data, performs additional statistical calculations, and returns p- and d-values which constitute the main body of the final heat maps (upper right and lower left triangular quadrants, respectively). Outputted files are tailored for subsequent import to GraphPad Prism.

Custom blank GraphPad Prism templates (formatted for compatibility with MATLAB outputs) as well as custom .groovy and .ijm scripts (used as a part of the QUINT pipeline) can be provided upon reasonable request.
