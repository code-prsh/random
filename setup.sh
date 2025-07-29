#!/bin/bash
mkdir -p ~/.streamlit/
echo "\n[server]\nheadless = true\nport = $PORT\nenableCORS = false\n\n[browser]\ngatherUsageStats = false\n\n[runner]\nfastReruns = true\n\n[logger]\nlevel = "info"\n\n[client]\nshowErrorDetails = true" > ~/.streamlit/config.toml
