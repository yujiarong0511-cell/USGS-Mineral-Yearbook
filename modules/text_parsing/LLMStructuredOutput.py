# Input: System Prompt + Local Prompt
# Output: A list of dictionaries in JSON-schema

from openai import OpenAI
from pydantic import BaseModel, Field
import textwrap
import json

# -----------------Alibaba Cloud API--------------------
def qwen_structured(input_text, 
                    base_url='http://localhost:11434/v1', api_key='ollama', model='qwen3:30b-a3b'):

    ## 1. Define response schema
    class Policy(BaseModel): # Properties to extract from each mention of policy. Clarify their (1) data type, (2) default value, (3) description/definition, (4) examples (should be wrapped up in a list)
        issuing_year: int | None = Field(default=None, description='The year when the law/regulation (and amendment) is enacted or the policy document is issued.')
        issuing_month: int | None = Field(default=None, description='The month when the law/regulation (and amendment) is enacted or the policy document is issued.') # Optional
        issuing_day: int | None = Field(default=None, description='The day when the law/regulation is enacted or the policy is issued.') # Optional
        document_name: str | None  = Field(default=None, description='The name of the law/regulation/policy document, with reference_id attached in square brackets', examples=['Mining Code [Decree-law (Act) No. 227]', 'Constitutional Amendments nos. 6'])
        reference_id: str | None = Field(default=None, description='The reference number/ID of the law/regulation/policy document.', examples=['Decree-law (Act) No. 227', 'Constitutional Amendments nos. 6'])
        policy_domain: list[str] | None = Field(default=None, description='The policy domain(s) the law/regulation/policy targeting at, which SHOULD BE selected from the following list: export_import_restriction/foreign_investment_access/equity_structure_state_ownership/technological_transfer/domestic_content_requirement/mining_concession/land_concession/tax_royalty/subsidy/labor_standard/environmental_standard/community_interest/foreign_exchange/geodata_regulation/comprehensive_framework', examples=[['export_import_restriction', 'domestic_content_requirement'],['comprehensive_framework']])
        legal_hierarchy: str | None = Field(default=None, description='The type of document is, which SHOULD BE extracted from the "reference_id" field.', examples=['law','constitution','decree','regulation','directive','notification'])
        issuing_authority: str | None = Field(default=None, description='The name of authority that enacted the measure, not applicable to constitution and its amendments, if mentioned in the input. Note it is NOT the enforcement authority.', examples=['national congress', 'Ministry of Foreign Commerce']) # Optional
        targeted_mineral: list[str] | None = Field(default=None, description='The mineral product(s) targeted by the law/regulation/policy, if mentioned in the input.', examples=[['nickel'],['cobalt','rare earth']]) # Optional
        is_local: bool | None = Field(default=None, description='Whether the law/regulation/policy is only applicable to part of the country RATHER THAN nation-wide law/regulation/policy.', examples=['True','False']) # Optional: administrative level (state/local)
        is_amendment: bool | None = Field(default=None, description='Whether the law/regulation/policy is an amendment to previous law/regulation/policy.', examples=['True','False']) # Optional

    class PolicyList(BaseModel): # We expect LLM to return a list of policies where each item satisfies the above schema
        policies: list[Policy]
    
    json_schema_string = json.dumps(PolicyList.model_json_schema()) # Convert defined response format into JSON schema
    
    ## 2. Design systemic prompt for overall guidance: (1) Role and output schema, (2) Task requirements, (3) Hints
    system_prompt = textwrap.dedent(f'''\
                                        [Your Role] Please parse the unstructured text into structured data. You must output a valid JSON object matching this exact structural schema: {json_schema_string}
                                        [Overall Requirements] There are indefinite numbers of laws/regulations/policies in each input. Find as many laws/regulations/policies as you can (but make sure not to include irrelevant contents), and analyze them one-by-one. But make sure ONLY laws/regulations/policies with clearly defined (1) year of enactment/issuing AND (2) reference number/ID (except for Constitution) can exist in your response. You MUST copy-and-paste the EXACT (1) year AND (2) reference number/ID mentioned from the input.
                                        [Notes] (1) Usually each paragraph from the input describes a policy issue or policy domain (but sometimes a policy issue or domain could span across several paragraphs). (2) Input often contains contents unrelated to laws/regulation/policies, such as introduction of the country's regulatory framework, annual statistical data among others, which should be omitted. (3) Each paragraph is usually ended with a string of in-text citation. DO NOT mistake them as laws/regulations/policies.
                                      ''')

    ## 3. Invoke LLM
    client = OpenAI(
        base_url = base_url, # API-URL ('/v1' suffix required)
        api_key = api_key # API-KEY
    )
    
    response = client.chat.completions.create(
        model = model, # Clarify the LLM to invoke
        # Impart prompts to LLM. Two parameters are compulsory: (1)'role' (whose value must be 'system' or 'user' or 'assistant'); (2)'content' (text of the prompt)
        messages=[
            {'role': 'system', 
             'content': system_prompt}, # Systemic prompt
            {'role': 'user', 
             'content': input_text} # Unstructured text input
            ],
        response_format={"type": "json_object"}, # Constrain the output format to JSON schema
        extra_body={
            'think': False # Prohibit invoking 'think' mode of the model
        }
    )

    extracted_data = response.choices[0].message.content
    return extracted_data


# -----------------Local Deployment (Ollama)-------------------
system_prompt_default = textwrap.dedent(f'''\
                                        [Your Role] Please parse the unstructured text into structured data. Your response should be in JSON schema (actually a list of analytical outputs in JSON schema).
                                        [Overall Requirements] There are indefinite numbers of laws/regulations/policies in each input. Find as many laws/regulations/policies as you can (but make sure not to include irrelevant contents), and analyze them one-by-one. But make sure ONLY laws/regulations/policies with clearly defined (1) year of enactment/issuing AND (2) reference number/ID (except for Constitution) can exist in your response. You MUST copy-and-paste the EXACT (1) year AND (2) reference number/ID mentioned from the input.
                                        [Notes] (1) Usually each paragraph from the input describes a policy issue or policy domain (but sometimes a policy issue or domain could span across several paragraphs). (2) Input often contains contents unrelated to laws/regulation/policies, such as introduction of the country's regulatory framework, annual statistical data among others, which should be omitted. (3) Each paragraph is usually ended with a string of in-text citation. DO NOT mistake them as laws/regulations/policies.
                                      ''')
def ollama_structured(input_text, system_prompt=system_prompt_default, 
                      base_url='http://localhost:11434/v1', api_key='ollama', model='qwen3:30b-a3b'):

    class Policy(BaseModel):
        issuing_year: int | None = Field(default=None, description='The year when the law/regulation (and amendment) is enacted or the policy document is issued.')
        issuing_month: int | None = Field(default=None, description='The month when the law/regulation (and amendment) is enacted or the policy document is issued.') # Optional
        issuing_day: int | None = Field(default=None, description='The day when the law/regulation is enacted or the policy is issued.') # Optional
        document_name: str | None  = Field(default=None, description='The name of the law/regulation/policy document, with reference_id attached in square brackets', examples=['Mining Code [Decree-law (Act) No. 227]', 'Constitutional Amendments nos. 6'])
        reference_id: str | None = Field(default=None, description='The reference number/ID of the law/regulation/policy document.', examples=['Decree-law (Act) No. 227', 'Constitutional Amendments nos. 6'])
        policy_domain: list[str] | None = Field(default=None, description='The policy domain(s) the law/regulation/policy targeting at, which SHOULD BE selected from the following list: export_import_restriction/foreign_investment_access/equity_structure_state_ownership/technological_transfer/domestic_content_requirement/mining_concession/land_concession/tax_royalty/subsidy/labor_standard/environmental_standard/community_interest/foreign_exchange/geodata_regulation/comprehensive_framework', examples=[['export_import_restriction', 'domestic_content_requirement'],['comprehensive_framework']])
        legal_hierarchy: str | None = Field(default=None, description='The type of document is, which SHOULD BE extracted from the "reference_id" field.', examples=['law','constitution','decree','regulation','directive','notification'])
        issuing_authority: str | None = Field(default=None, description='The name of authority that enacted the measure, not applicable to constitution and its amendments, if mentioned in the input. Note it is NOT the enforcement authority.', examples=['national congress', 'Ministry of Foreign Commerce']) # Optional
        targeted_mineral: list[str] | None = Field(default=None, description='The mineral product(s) targeted by the law/regulation/policy, if mentioned in the input.', examples=[['nickel'],['cobalt','rare earth']]) # Optional
        is_local: bool | None = Field(default=None, description='Whether the law/regulation/policy is only applicable to part of the country RATHER THAN nation-wide law/regulation/policy.', examples=['True','False']) # Optional: administrative level (state/local)
        is_amendment: bool | None = Field(default=None, description='Whether the law/regulation/policy is an amendment to previous law/regulation/policy.', examples=['True','False']) # Optional
    class PolicyList(BaseModel):
        policies: list[Policy]
    
    client = OpenAI(
        base_url = base_url,
        api_key = api_key
    )

    response = client.beta.chat.completions.parse(
        model = model, # Clarify the LLM to invoke
        # Impart the prompt to LLM. Two parameters are compulsory: (1)'role' (whose value must be 'system' or 'user'); (2)'content' (text of the prompt)
        messages=[
            {'role': 'system', 
             'content': system_prompt},
            {'role': 'user', 
             'content': input_text}
            ],
        response_format=PolicyList, # Constrain the output format to JSON schema
        extra_body={
            'think': False # Prohibit invoking 'think' mode of the model
        }
    )

    extracted_data = response.choices[0].message.parsed
    return extracted_data