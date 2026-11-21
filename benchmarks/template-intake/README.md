# Reproduce the three-template intake trial

This is a secondary experiment designed after the campus audit tied. Its decision rules were published before model runs. Only the five files in the solver's `inputs/` directory were shared with the model. The protocol, score code and other runs were outside its working directory. [Registered protocol](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/blob/dd5e5ff9e6c3d20d6f9472ea9547bf7f3fd1d535/benchmarks/template-intake/PROTOCOL.md) · [Results](../../docs/intake.md)

## Score the actual delivered files

Download the synthetic evidence ZIP from the release page and check its SHA-256 against the release checksum. Extract it before running the scorer. Use the folder containing the five supplied intake files as INPUTS and the folder containing the three DOCX files for one run as OUTPUT. Run this for each of the six outputs. Python 3.10 or later is sufficient. [Release attachments](https://github.com/Spark-xY6/dsh-plugin-bio-thesis/releases/tag/v0.3.0)

```sh
python -B -X utf8 benchmarks/template-intake/controls.py INPUTS
python -B -X utf8 benchmarks/template-intake/judge.py INPUTS OUTPUT
```

Read the four booleans for every file as well as the total. The first checks whether the delivered package is an editable DOCX. The other three test the required field values, anonymity and preservation of the original template. The control suite contains 18 positive, negative and equivalent-serialization cases. Whole-file hashes identify the actual returned files. ZIP timestamps or XML serialization can change the whole-file digest without changing the document's score.

## Build the synthetic templates

From a clean checkout, build the campus fixtures with Node first. Run the intake builder with a fresh output directory and copy `TASK.md` beside the four files it creates. The frozen corpus was built on Windows with Python 3.13.9. Other ZIP implementations and newline settings can produce different bytes. Use the evidence ZIP for exact-byte replay.

```powershell
node benchmarks/campus-v1/build.mjs
python -B -X utf8 benchmarks/template-intake/build.py benchmarks/campus-v1/inputs/school.dotx benchmarks/template-intake/inputs
Copy-Item benchmarks/template-intake/TASK.md benchmarks/template-intake/inputs/TASK.md
python -B -X utf8 benchmarks/template-intake/controls.py benchmarks/template-intake/inputs
```

## Repeat the live model runs

The published `run.mjs` is the Windows driver used in the experiment. It verifies the manifest and creates six fresh workspaces. Supply the frozen five inputs in `benchmarks/template-intake/inputs`, ensure the v0.2.0 commit is available locally and start with no existing `trials` directory in that checkout. Use a Windows Codex CLI 0.169.0 with a working workspace-write sandbox, a signed-in account with access to the specified model and an explicit Python executable. The model service requires network access. Document-processing commands in the solver sandbox have networking disabled. Rerunning six model sessions can incur usage charges.

```powershell
$env:THESIS_GIT_ROOT = (Get-Location).Path
$env:THESIS_PYTHON = (Get-Command python -CommandType Application | Select-Object -First 1).Source
$env:CODEX_EXECUTABLE = (Get-Command codex.exe -CommandType Application | Select-Object -First 1).Source
node benchmarks/template-intake/run.mjs
```

If the CLI launcher uses a wrapper, set `CODEX_EXECUTABLE` to the native executable. Scoring the fixed evidence does not require model access. A new model run is stochastic. Keep any new raw logs local, since the host may record account details and absolute paths.
