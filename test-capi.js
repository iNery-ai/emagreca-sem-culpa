const CAPI = require('./capi');

async function runTests() {
  console.log('--- RUNNING META CAPI PAYLOAD GENERATOR TESTS ---\n');

  let passed = 0;
  let failed = 0;

  function assert(condition, message) {
    if (condition) {
      console.log(`✅ [PASS] ${message}`);
      passed++;
    } else {
      console.error(`❌ [FAIL] ${message}`);
      failed++;
    }
  }

  // Test Case 1: Sanitization (omitting null / empty / [null] values)
  try {
    const rawUserData = {
      em: 'TestEmail@Example.com', // should be hashed and lowercase
      ph: [null],                  // should be omitted
      fn: null,                    // should be omitted
      ln: '',                      // should be omitted
      ct: ['São Paulo', ''],       // should filter out empty string, keep São Paulo hashed
      country: [],                 // empty array should be omitted
      client_ip_address: '127.0.0.1' // non-sensitive, should stay unhashed as scalar
    };

    const customData = {
      value: '9.90',
      currency: 'brl',
      content_name: 'Guia Emagrecimento'
    };

    const payload = await CAPI.generateMetaCAPIPayload('Purchase', rawUserData, customData);
    const event = payload.data[0];

    // Assert event time is dynamic Unix timestamp (close to now)
    const now = Math.floor(Date.now() / 1000);
    assert(Math.abs(event.event_time - now) < 5, `event_time should be a recent Unix Timestamp (${event.event_time})`);

    // Assert action source
    assert(event.action_source === 'website', `action_source defaults to "website"`);

    // Assert user_data sanitization & hashing
    assert(event.user_data.ph === undefined, 'user_data.ph containing [null] must be programmatically omitted');
    assert(event.user_data.fn === undefined, 'user_data.fn being null must be programmatically omitted');
    assert(event.user_data.ln === undefined, 'user_data.ln being empty string must be programmatically omitted');
    assert(event.user_data.country === undefined, 'user_data.country being empty array must be programmatically omitted');

    // Assert correct hashing of em (lowercase and hashed)
    // "testemail@example.com" hashed is "91137f03fb9029daca900eef79941c05cdb2271ca076f2624c6f0362e58c2a96"
    const expectedEmailHash = '91137f03fb9029daca900eef79941c05cdb2271ca076f2624c6f0362e58c2a96';
    assert(Array.isArray(event.user_data.em), 'user_data.em must be an array');
    assert(event.user_data.em[0] === expectedEmailHash, `user_data.em[0] should be correctly hashed to lowercase: ${event.user_data.em[0]}`);

    // Assert correct hashing of ct ("são paulo" hashed)
    // "são paulo" hashed is "577abdbf90dadd651458eee7576c6e3684b5c27beabd465cc4bb3c42441b5b38"
    const expectedCtHash = '577abdbf90dadd651458eee7576c6e3684b5c27beabd465cc4bb3c42441b5b38';
    assert(event.user_data.ct.length === 1, 'user_data.ct should filter out empty strings and have length 1');
    assert(event.user_data.ct[0] === expectedCtHash, 'user_data.ct[0] should be correctly hashed');

    // Assert client_ip_address remains unhashed and is a scalar string
    assert(event.user_data.client_ip_address === '127.0.0.1', 'Non-sensitive client_ip_address must remain unhashed');

    // Assert custom_data format
    assert(event.custom_data.value === 9.9, `custom_data.value must be parsed as a float: ${event.custom_data.value}`);
    assert(event.custom_data.currency === 'BRL', `custom_data.currency must be uppercase "BRL": ${event.custom_data.currency}`);
    assert(event.custom_data.content_name === 'Guia Emagrecimento', 'Other custom_data fields should be preserved');

  } catch (err) {
    console.error('Test Case 1 failed with error:', err);
    failed++;
  }

  // Test Case 2: Pre-hashed sensitive fields should not be double-hashed
  try {
    const rawUserData = {
      // Already hashed e-mail (from the user's prompt example)
      em: '7b17fb0bd173f625b58636fb796407c22b3d16fc78302d79f0fd30c2fc2fc068',
      ph: ['+5511999999999'] // raw phone
    };

    const payload = await CAPI.generateMetaCAPIPayload('Purchase', rawUserData, { value: 9.9 });
    const event = payload.data[0];

    // em should not change because it's already a valid SHA-256 hash
    assert(
      event.user_data.em[0] === '7b17fb0bd173f625b58636fb796407c22b3d16fc78302d79f0fd30c2fc2fc068',
      'Pre-hashed user data must not be double-hashed'
    );

    // ph should be hashed
    const expectedPhoneHash = await CAPI.sha256('+5511999999999');
    assert(
      event.user_data.ph[0] === expectedPhoneHash,
      'Raw user phone must be hashed'
    );
  } catch (err) {
    console.error('Test Case 2 failed with error:', err);
    failed++;
  }

  console.log(`\nSummary: ${passed} passed, ${failed} failed`);
  if (failed > 0) {
    process.exit(1);
  } else {
    console.log('All tests completed successfully!');
  }
}

runTests();
