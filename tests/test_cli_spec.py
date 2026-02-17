"""Tests for CLI spec file loading and batch generation."""

import yaml
from typer.testing import CliRunner

from docxfix.cli import app

runner = CliRunner()


def test_create_with_spec_yaml(tmp_path):
    """Test create command with YAML spec file."""
    # Create a simple spec file
    spec_file = tmp_path / "test.yaml"
    spec_data = {
        "title": "Test Document",
        "author": "Test Author",
        "paragraphs": [
            {"text": "First paragraph"},
            {"text": "Second paragraph"},
        ],
    }
    with open(spec_file, "w") as f:
        yaml.dump(spec_data, f)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file)]
    )

    assert result.exit_code == 0
    assert "✓ Docx fixture created successfully!" in result.stdout
    assert output_file.exists()


def test_create_with_spec_json(tmp_path):
    """Test create command with JSON spec file."""
    # Create a simple spec file
    spec_file = tmp_path / "test.json"
    spec_data = """
    {
        "title": "JSON Test",
        "author": "JSON Author",
        "paragraphs": [
            {"text": "JSON paragraph"}
        ]
    }
    """
    with open(spec_file, "w") as f:
        f.write(spec_data)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file)]
    )

    assert result.exit_code == 0
    assert "✓ Docx fixture created successfully!" in result.stdout
    assert output_file.exists()


def test_create_with_spec_verbose(tmp_path):
    """Test create command with spec file in verbose mode."""
    spec_file = tmp_path / "test.yaml"
    spec_data = {
        "title": "Verbose Test",
        "paragraphs": [{"text": "Test paragraph"}],
    }
    with open(spec_file, "w") as f:
        yaml.dump(spec_data, f)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file), "--verbose"]
    )

    assert result.exit_code == 0
    assert "Creating docx fixture at:" in result.stdout
    assert "Using spec file:" in result.stdout
    assert "Parsing spec file:" in result.stdout
    assert "Document generated" in result.stdout
    assert "Validating document..." in result.stdout
    assert "Validation passed" in result.stdout


def test_create_with_invalid_spec(tmp_path):
    """Test create command with invalid spec file."""
    spec_file = tmp_path / "invalid.yaml"
    spec_data = {
        "title": "Invalid",
        # Missing required 'paragraphs' field
    }
    with open(spec_file, "w") as f:
        yaml.dump(spec_data, f)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file)]
    )

    assert result.exit_code == 1
    # Error message appears in output
    combined_output = result.stdout + result.output
    assert "Spec parsing error" in combined_output or "paragraphs" in combined_output


def test_create_with_missing_spec_file(tmp_path):
    """Test create command with non-existent spec file."""
    spec_file = tmp_path / "nonexistent.yaml"
    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file)]
    )

    assert result.exit_code == 1
    # Error message appears in output
    combined_output = result.stdout + result.output
    assert "File not found" in combined_output or "not found" in combined_output.lower()


def test_create_with_spec_no_validate(tmp_path):
    """Test create command with spec file and validation disabled."""
    spec_file = tmp_path / "test.yaml"
    spec_data = {
        "title": "No Validate Test",
        "paragraphs": [{"text": "Test"}],
    }
    with open(spec_file, "w") as f:
        yaml.dump(spec_data, f)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app,
        ["create", str(output_file), "--spec", str(spec_file), "--no-validate"],
    )

    assert result.exit_code == 0
    assert "✓ Docx fixture created successfully!" in result.stdout
    assert output_file.exists()


def test_create_with_complex_spec(tmp_path):
    """Test create command with a complex spec including all features."""
    spec_file = tmp_path / "complex.yaml"
    spec_data = {
        "title": "Complex Document",
        "author": "Complex Author",
        "seed": 42,
        "paragraphs": [
            {"text": "Plain paragraph"},
            {
                "text": "Paragraph with tracked changes",
                "tracked_changes": [
                    {
                        "change_type": "insertion",
                        "text": "inserted text",
                        "author": "Editor",
                    }
                ],
            },
            {
                "text": "Paragraph with anchor text comment",
                "comments": [
                    {
                        "text": "This is a comment",
                        "anchor_text": "anchor text",
                        "author": "Reviewer",
                    }
                ],
            },
            {
                "text": "Numbered item at level 1",
                "numbering": {"level": 1},
            },
        ],
    }
    with open(spec_file, "w") as f:
        yaml.dump(spec_data, f)

    output_file = tmp_path / "output.docx"

    result = runner.invoke(
        app, ["create", str(output_file), "--spec", str(spec_file)]
    )

    assert result.exit_code == 0
    assert output_file.exists()


def test_batch_simple(tmp_path):
    """Test batch command with simple manifest."""
    # Create spec files
    spec1 = tmp_path / "spec1.yaml"
    spec1_data = {
        "title": "Document 1",
        "paragraphs": [{"text": "Content 1"}],
    }
    with open(spec1, "w") as f:
        yaml.dump(spec1_data, f)

    spec2 = tmp_path / "spec2.yaml"
    spec2_data = {
        "title": "Document 2",
        "paragraphs": [{"text": "Content 2"}],
    }
    with open(spec2, "w") as f:
        yaml.dump(spec2_data, f)

    # Create manifest
    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "spec1.yaml", "output": "doc1.docx"},
            {"id": "doc2", "spec": "spec2.yaml", "output": "doc2.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 0
    assert "Batch generation complete" in result.stdout
    assert "Success: 2" in result.stdout
    assert "Failed: 0" in result.stdout
    assert (out_dir / "doc1.docx").exists()
    assert (out_dir / "doc2.docx").exists()


def test_batch_verbose(tmp_path):
    """Test batch command in verbose mode."""
    spec1 = tmp_path / "spec1.yaml"
    spec1_data = {
        "title": "Document 1",
        "paragraphs": [{"text": "Content"}],
    }
    with open(spec1, "w") as f:
        yaml.dump(spec1_data, f)

    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "spec1.yaml", "output": "doc1.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app,
        [
            "batch",
            "--manifest",
            str(manifest_file),
            "--out-dir",
            str(out_dir),
            "--verbose",
        ],
    )

    assert result.exit_code == 0
    assert "Loading batch manifest" in result.stdout
    assert "Output directory" in result.stdout
    assert "Processing 'doc1'" in result.stdout
    assert "Spec:" in result.stdout
    assert "Output:" in result.stdout
    assert "Generated successfully" in result.stdout


def test_batch_with_invalid_spec(tmp_path):
    """Test batch command with one invalid spec."""
    # Valid spec
    spec1 = tmp_path / "spec1.yaml"
    spec1_data = {
        "title": "Document 1",
        "paragraphs": [{"text": "Content 1"}],
    }
    with open(spec1, "w") as f:
        yaml.dump(spec1_data, f)

    # Invalid spec
    spec2 = tmp_path / "spec2.yaml"
    spec2_data = {
        "title": "Document 2",
        # Missing paragraphs
    }
    with open(spec2, "w") as f:
        yaml.dump(spec2_data, f)

    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "spec1.yaml", "output": "doc1.docx"},
            {"id": "doc2", "spec": "spec2.yaml", "output": "doc2.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 1
    assert "Success: 1" in result.stdout
    assert "Failed: 1" in result.stdout
    assert "Failed fixtures:" in result.stdout
    assert "doc2" in result.stdout
    assert (out_dir / "doc1.docx").exists()
    assert not (out_dir / "doc2.docx").exists()


def test_batch_with_missing_spec_file(tmp_path):
    """Test batch command with missing spec file."""
    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "nonexistent.yaml", "output": "doc1.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 1
    assert "Failed: 1" in result.stdout
    assert "File not found" in result.stdout


def test_batch_with_missing_manifest(tmp_path):
    """Test batch command with non-existent manifest."""
    manifest_file = tmp_path / "nonexistent.yaml"
    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 1
    # Error message appears in output
    combined_output = result.stdout + result.output
    assert "Manifest error" in combined_output or "not found" in combined_output.lower()


def test_batch_with_invalid_manifest_format(tmp_path):
    """Test batch command with invalid manifest format."""
    manifest_file = tmp_path / "manifest.yaml"
    # Missing 'fixtures' key
    manifest_data = {"wrong_key": []}
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 1
    # Error message appears in output
    combined_output = result.stdout + result.output
    assert "Manifest error" in combined_output or "fixtures" in combined_output


def test_batch_with_missing_fixture_fields(tmp_path):
    """Test batch command with fixtures missing required fields."""
    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "spec.yaml"},  # Missing 'output'
            {"id": "doc2", "output": "doc2.docx"},  # Missing 'spec'
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 1
    assert "Failed: 2" in result.stdout
    assert "Missing 'output' field" in result.stdout
    assert "Missing 'spec' field" in result.stdout


def test_batch_no_validate(tmp_path):
    """Test batch command with validation disabled."""
    spec1 = tmp_path / "spec1.yaml"
    spec1_data = {
        "title": "Document 1",
        "paragraphs": [{"text": "Content"}],
    }
    with open(spec1, "w") as f:
        yaml.dump(spec1_data, f)

    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "spec1.yaml", "output": "doc1.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app,
        [
            "batch",
            "--manifest",
            str(manifest_file),
            "--out-dir",
            str(out_dir),
            "--no-validate",
        ],
    )

    assert result.exit_code == 0
    assert (out_dir / "doc1.docx").exists()


def test_batch_many_fixtures(tmp_path):
    """Test batch command with 20+ fixtures to meet acceptance criteria."""
    fixtures = []
    for i in range(25):
        spec_file = tmp_path / f"spec{i}.yaml"
        spec_data = {
            "title": f"Document {i}",
            "paragraphs": [{"text": f"Content {i}"}],
        }
        with open(spec_file, "w") as f:
            yaml.dump(spec_data, f)

        fixtures.append({
            "id": f"doc{i}",
            "spec": f"spec{i}.yaml",
            "output": f"doc{i}.docx",
        })

    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {"fixtures": fixtures}
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 0
    assert "Total: 25" in result.stdout
    assert "Success: 25" in result.stdout
    assert "Failed: 0" in result.stdout

    # Verify all files were created
    for i in range(25):
        assert (out_dir / f"doc{i}.docx").exists()


def test_batch_spec_paths_relative_to_manifest(tmp_path):
    """Test that spec paths in manifest are resolved relative to manifest directory."""
    # Create subdirectory for specs
    specs_dir = tmp_path / "specs"
    specs_dir.mkdir()

    spec1 = specs_dir / "spec1.yaml"
    spec1_data = {
        "title": "Document 1",
        "paragraphs": [{"text": "Content"}],
    }
    with open(spec1, "w") as f:
        yaml.dump(spec1_data, f)

    # Manifest in parent directory
    manifest_file = tmp_path / "manifest.yaml"
    manifest_data = {
        "fixtures": [
            {"id": "doc1", "spec": "specs/spec1.yaml", "output": "doc1.docx"},
        ]
    }
    with open(manifest_file, "w") as f:
        yaml.dump(manifest_data, f)

    out_dir = tmp_path / "output"

    result = runner.invoke(
        app, ["batch", "--manifest", str(manifest_file), "--out-dir", str(out_dir)]
    )

    assert result.exit_code == 0
    assert (out_dir / "doc1.docx").exists()


def test_create_without_spec_still_works(tmp_path):
    """Test that create command without spec still generates default document."""
    output_file = tmp_path / "output.docx"

    result = runner.invoke(app, ["create", str(output_file)])

    assert result.exit_code == 0
    assert output_file.exists()
