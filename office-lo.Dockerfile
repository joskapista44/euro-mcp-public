# LibreOffice headless engine for the office MCP.
#
# Built here rather than pulled from a published LibreOffice image on purpose. This
# container runs on the fleet host with a bind mount onto files the agents are working on -
# client material included - so whatever is inside it is running next to data that must not
# leak. A third-party image is opaque and can change under a moving tag; a Dockerfile is
# reviewable, pinned to what we actually install, and reproducible.
#
# Build:
#   docker build -f office-lo.Dockerfile -t euro-office-lo:1 .
#
# Run (one short-lived container per job, no network, single writable mount):
#   docker run --rm --network=none \
#     -v <jobdir>:/work -u $(id -u):$(id -g) -e HOME=/work euro-office-lo:1 \
#     soffice --headless --norestore \
#       -env:UserInstallation=file:///work/profile \
#       --convert-to xlsx --outdir /work/out /work/in.xlsx
#
# HOME=/work is not optional either. The uid passed with -u has no entry in the image's
# /etc/passwd, so LibreOffice cannot look up a home directory and falls back to somewhere
# it may not be able to write. Pointing HOME at the one writable mount avoids a failure
# that surfaces as a profile-init error rather than as anything about permissions.
#
# --network=none is not decoration: a spreadsheet can carry external references, and the
# converter never needs to reach anything. Without it, opening a hostile document is an
# outbound request we did not intend.
#
# The recalculation setting is NOT baked into the image. It goes into the per-job profile
# under /work, written by the caller - see office_recalc_gate.py (prepare_profile). Baking
# it in would hide the single setting the whole correctness of this tool depends on, and a
# later image rebuild that dropped it would not fail: it would quietly return stale numbers.

FROM debian:12-slim

# libreoffice-core alone cannot open xlsx; calc is the one that matters here, writer and
# impress are for the docx/pptx work that follows. No recommends, so no GUI stack, no Java:
# the conversion path does not use them and they would roughly double the image.
# The fonts are not optional - without them PDF output silently falls back to substitutes
# and the layout drifts.
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        libreoffice-core \
        libreoffice-calc \
        libreoffice-writer \
        libreoffice-impress \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Everything the job touches lives here and is bind-mounted at run time: the input, the
# output directory, and the throwaway LibreOffice profile. Nothing is written outside it.
WORKDIR /work

# No USER line on purpose. The caller passes -u <uid>:<gid> so the container writes as the
# invoking user and the produced files are owned by them rather than by root. A baked-in
# user would collide with that: soffice needs to write the profile inside the mount, and a
# uid mismatch there fails with a permission error that reads like a LibreOffice bug.

# No ENTRYPOINT either - the caller passes the full soffice command, which keeps the
# invocation visible at the call site instead of split between here and the code.
