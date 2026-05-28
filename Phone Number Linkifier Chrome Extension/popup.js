const patternsForm = document.getElementById('patterns');
const resetButton = document.getElementById('reset');
const statusElement = document.getElementById('status');

let statusTimer = null;

function showStatus(message) {
  statusElement.textContent = message;
  if (statusTimer) clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    statusElement.textContent = '';
  }, 1800);
}

function getFormSettings() {
  const settings = {};
  for (const pattern of PhonePatternSettings.PATTERN_DEFINITIONS) {
    const input = patternsForm.elements[pattern.id];
    settings[pattern.id] = Boolean(input && input.checked);
  }
  return settings;
}

async function saveFormSettings() {
  try {
    await PhonePatternSettings.saveSettings(getFormSettings());
    showStatus('Saved');
  } catch (err) {
    console.error('Could not save phone pattern settings:', err);
    showStatus('Could not save settings');
  }
}

function renderPatterns(settings) {
  patternsForm.textContent = '';

  for (const pattern of PhonePatternSettings.PATTERN_DEFINITIONS) {
    const label = document.createElement('label');
    label.className = 'pattern';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.name = pattern.id;
    checkbox.checked = Boolean(settings[pattern.id]);

    const details = document.createElement('span');

    const title = document.createElement('span');
    title.className = 'pattern-title';

    const name = document.createElement('span');
    name.className = 'pattern-name';
    name.textContent = pattern.label;

    const example = document.createElement('span');
    example.className = 'pattern-example';
    example.textContent = pattern.example;

    const description = document.createElement('p');
    description.className = 'pattern-description';
    description.textContent = pattern.description;

    title.append(name, example);
    details.append(title, description);
    label.append(checkbox, details);
    patternsForm.append(label);
  }
}

patternsForm.addEventListener('change', saveFormSettings);

resetButton.addEventListener('click', async () => {
  const defaults = PhonePatternSettings.getDefaultSettings();
  renderPatterns(defaults);
  try {
    await PhonePatternSettings.saveSettings(defaults);
    showStatus('Defaults restored');
  } catch (err) {
    console.error('Could not reset phone pattern settings:', err);
    showStatus('Could not reset settings');
  }
});

PhonePatternSettings.loadSettings()
  .then(renderPatterns)
  .catch((err) => {
    console.error('Could not load phone pattern settings:', err);
    renderPatterns(PhonePatternSettings.getDefaultSettings());
    showStatus('Using defaults');
  });
