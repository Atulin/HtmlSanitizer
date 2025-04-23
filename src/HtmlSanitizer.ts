const _parser = new DOMParser();

export class HtmlSanitizer {
	public tagWhitelist: Record<string, boolean> = {
		A: true,
		ABBR: true,
		B: true,
		BLOCKQUOTE: true,
		BODY: true,
		BR: true,
		CENTER: true,
		CODE: true,
		DD: true,
		DIV: true,
		DL: true,
		DT: true,
		EM: true,
		FONT: true,
		H1: true,
		H2: true,
		H3: true,
		H4: true,
		H5: true,
		H6: true,
		HR: true,
		I: true,
		IMG: true,
		LABEL: true,
		LI: true,
		OL: true,
		P: true,
		PRE: true,
		SMALL: true,
		SOURCE: true,
		SPAN: true,
		STRONG: true,
		SUB: true,
		SUP: true,
		TABLE: true,
		TBODY: true,
		TR: true,
		TD: true,
		TH: true,
		THEAD: true,
		UL: true,
		U: true,
		VIDEO: true,
	};

	private contentTagWhiteList: Record<string, boolean> = {
		FORM: true,
		"GOOGLE-SHEETS-HTML-ORIGIN": true,
	}; //tags that will be converted to DIVs

	public attributeWhitelist: Record<string, boolean> = {
		align: true,
		color: true,
		controls: true,
		height: true,
		href: true,
		id: true,
		src: true,
		style: true,
		target: true,
		title: true,
		type: true,
		width: true,
	};

	public cssWhitelist: Record<string, boolean> = {
		"background-color": true,
		color: true,
		"font-size": true,
		"font-weight": true,
		"text-align": true,
		"text-decoration": true,
		width: true,
	};

	public schemaWhiteList: string[] = [
		"http:",
		"https:",
		"data:",
		"m-files:",
		"file:",
		"ftp:",
		"mailto:",
		"pw:",
	]; //which "protocols" are allowed in "href", "src" etc

	private uriAttributes: Record<string, boolean> = {
		href: true,
		action: true,
	};

	public SanitizeHtml = (
		htmlString: string,
		extraSelector?: string,
	): string => {
		let input = htmlString.trim();
		if (input === "") {
			return "";
		} //to save performance

		//firefox "bogus node" workaround for wysiwyg's
		if (input === "<br>") {
			return "";
		}

		if (input.indexOf("<body") === -1) {
			input = `<body>${input}</body>`;
		} //add "body" otherwise some tags are skipped, like <style>

		const doc = _parser.parseFromString(input, "text/html");

		//DOM clobbering check (damn you firefox)
		if (doc.body.tagName !== "BODY") {
			doc.body.remove();
		}

		if (typeof doc.createElement !== "function") {
			(doc.createElement as unknown as HTMLElement).remove();
		}

		const makeSanitizedCopy = (node: HTMLElement | Node): Node => {
			let newNode: HTMLElement | Node;
			if (node.nodeType === Node.TEXT_NODE) {
				newNode = node.cloneNode(true);
			} else if (
				node.nodeType === Node.ELEMENT_NODE &&
				node instanceof HTMLElement &&
				(this.tagWhitelist[node.tagName] ||
					this.contentTagWhiteList[node.tagName] ||
					(extraSelector && node.matches(extraSelector)))
			) {
				//is tag allowed?
				if (this.contentTagWhiteList[node.tagName]) {
					newNode = doc.createElement("DIV");
				} //convert to DIV
				else {
					newNode = doc.createElement(node.tagName);
				}

				for (let i = 0; i < node.attributes.length; i++) {
					const attr = node.attributes[i];
					if (
						this.attributeWhitelist[attr.name] &&
						!(attr.value in doc)
					) {
						if (attr.name === "style") {
							for (let s = 0; s < node.style.length; s++) {
								const styleName = node.style[s];
								if (
									this.cssWhitelist[styleName] &&
									newNode instanceof HTMLElement
								) {
									newNode.style.setProperty(
										styleName,
										node.style.getPropertyValue(styleName),
									);
								}
							}
						} else if (newNode instanceof HTMLElement) {
							if (this.uriAttributes[attr.name]) {
								//if this is an "uri" attribute, that can have "JavaScript:" or something
								if (
									attr.value.indexOf(":") > -1 &&
									!this.startsWithAny(
										attr.value,
										this.schemaWhiteList,
									)
								) {
									continue;
								}
							}
							newNode.setAttribute(attr.name, attr.value);
						}
					}
				}

				for (let i = 0; i < node.childNodes.length; i++) {
					const subCopy = makeSanitizedCopy(node.childNodes[i]);
					newNode.appendChild(subCopy);
				}

				//remove useless empty spans (lots of those when pasting from MS Outlook)
				if (
					newNode instanceof HTMLElement &&
					(newNode.tagName === "SPAN" ||
						newNode.tagName === "B" ||
						newNode.tagName === "I" ||
						newNode.tagName === "U") &&
					newNode.innerHTML.trim() === ""
				) {
					return doc.createDocumentFragment();
				}
			} else {
				newNode = doc.createDocumentFragment();
			}
			return newNode;
		};

		const resultElement = makeSanitizedCopy(doc.body);

		return (resultElement as HTMLElement).innerHTML.replace(
			/div><div/g,
			"div>\n<div",
		); //replace is just for cleaner code
	};

	private startsWithAny(str: string, substrings: string[]): boolean {
		for (let i = 0; i < substrings.length; i++) {
			if (str.indexOf(substrings[i]) === 0) {
				return true;
			}
		}
		return false;
	}
}
