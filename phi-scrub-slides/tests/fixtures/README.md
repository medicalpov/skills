Synthetic test decks are generated, not committed (the repo ignores all .pptx files so a real deck can
never be committed by accident).

    python3 tests/make_fixture.py                      # -> tests/fixtures/synthetic_case.pptx
    python3 tests/make_fixture.py tests/fixtures/x.pptm  # macro-enabled variant

Every identifier in them is fake ("Jane Testpatient", MRN 00012345, DOB 01/02/1950, 555-0100,
*.example.com, a made-up GPS position).
