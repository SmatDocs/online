/* -*- js-indent-level: 8 -*- */
/*
 * Copyright the Collabora Online contributors.
 *
 * SPDX-License-Identifier: MPL-2.0
 */

describe('CallPythonScriptBridge', function() {
	const bridge = (window as any).CallPythonScriptBridge;
	const commandName = 'vnd.sun.star.script:ApplyWordActionPlan.py$applyWordActionPlan?language=Python&location=share';

	it('keeps the Python callback pending through unrelated UNO results', function() {
		const source = { postMessage() {} };
		const requestedAt = 1000;
		let requests = bridge.registerRequest([], {
			commandName,
			requestedAt,
			source,
		}, requestedAt);

		for (const unrelatedCommand of ['.uno:DeleteBookmark', '.uno:Save', '.uno:Refresh']) {
			const unmatched = bridge.takeMatchingResult(
				requests,
				unrelatedCommand,
				{ commandName: unrelatedCommand, success: true },
				requestedAt + 1);
			nodeassert.strictEqual(unmatched.completedRequest, null);
			nodeassert.strictEqual(unmatched.responseMessage, null);
			nodeassert.strictEqual(unmatched.remainingRequests.length, 1);
			requests = unmatched.remainingRequests;
		}

		const scriptResult = {
			commandName,
			success: true,
			result: {
				type: 'string',
				value: JSON.stringify({
					event: 'word_action_plan_applied',
					plan_id: 'wp_test',
					file_id: 'file_test',
				}),
			},
		};
		const matched = bridge.takeMatchingResult(
			requests, commandName, scriptResult, requestedAt + 10);

		nodeassert.strictEqual(matched.completedRequest.source, source);
		nodeassert.strictEqual(matched.remainingRequests.length, 0);
		nodeassert.deepStrictEqual(matched.responseMessage, {
			MessageId: 'CallPythonScript-Result',
			SendTime: requestedAt + 10,
			Values: scriptResult,
		});
	});
});
