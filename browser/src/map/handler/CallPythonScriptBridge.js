/* -*- js-indent-level: 8 -*- */
/*
 * Copyright the Collabora Online contributors.
 *
 * SPDX-License-Identifier: MPL-2.0
 */

window.CallPythonScriptBridge = {
	requestMaxAgeMs: 120000,

	registerRequest: function (requests, request, now) {
		var pending = Array.isArray(requests) ? requests : [];
		pending = pending.filter(function (existingRequest) {
			return now - existingRequest.requestedAt < this.requestMaxAgeMs;
		}, this);
		pending.push(request);
		return pending;
	},

	takeMatchingResult: function (requests, commandName, result, sendTime) {
		var pending = Array.isArray(requests) ? requests : [];
		var requestIndex = pending.findIndex(function (request) {
			return request.commandName === commandName;
		});
		if (requestIndex === -1) {
			return {
				completedRequest: null,
				remainingRequests: pending,
				responseMessage: null,
			};
		}

		var remainingRequests = pending.slice();
		var completedRequest = remainingRequests.splice(requestIndex, 1)[0];
		return {
			completedRequest: completedRequest,
			remainingRequests: remainingRequests,
			responseMessage: {
				MessageId: 'CallPythonScript-Result',
				SendTime: sendTime,
				Values: result,
			},
		};
	},
};
