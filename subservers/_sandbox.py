from collections.abc import Callable

from pydantic_monty import CollectString, Monty, MontyError, ResourceLimits

_LIMITS = ResourceLimits(
    max_duration_secs=10.0,
    max_memory=128 * 1024 * 1024,
)


async def run_sandboxed(
    code: str,
    functions: dict[str, Callable],
) -> str:

    collected = CollectString()

    try:

        monty = await Monty.acreate(code)

        result = await monty.run_async(
            limits=_LIMITS,
            external_functions=functions,
            print_callback=collected,
        )

    except MontyError as error:
        raise ValueError(f"Script failed: {error}") from error

    output = collected.output

    if result is not None:
        output += f"{result!r}\n"

    return output
