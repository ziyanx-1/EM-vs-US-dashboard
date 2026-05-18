#!/bin/bash
set -e
streamlit run mas_analytics.py \
    --server.port 9999 \
    --server.address 0.0.0.0 \
    --server.headless true
